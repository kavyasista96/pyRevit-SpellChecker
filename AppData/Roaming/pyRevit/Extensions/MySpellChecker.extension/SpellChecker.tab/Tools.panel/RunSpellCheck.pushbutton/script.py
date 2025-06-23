# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, TextNote, ViewSchedule, Transaction
from pyrevit import forms, script

output = script.get_output()
pending_edits = []  # (element, new_text, cell_pos)

def ask_fix(word, context):
    prompt = """
Word: '{}'
Context: {}
Type replacement OR type:
- 'skip' to ignore
- 'stop' to save and exit spell check
""".format(word, context)
    return forms.ask_for_string(default=word, prompt=prompt)

def check_spelling(text, context, element=None, cell_pos=None):
    global pending_edits
    original_text = text
    words = text.split()
    updated_text = text

    for word in words:
        fixed = ask_fix(word, context)

        if not fixed:
            continue
        if fixed.lower() == "skip":
            continue

        if fixed.lower() == "stop":
            if updated_text != original_text:
                pending_edits.append((element, updated_text, cell_pos))
            return "STOP_COMMAND"

        updated_text = updated_text.replace(word, fixed)

    if updated_text != original_text:
        pending_edits.append((element, updated_text, cell_pos))

def commit_changes():
    doc = __revit__.ActiveUIDocument.Document
    t = Transaction(doc, "Apply Spell Check Fixes")
    t.Start()
    for elem, new_text, cell in pending_edits:
        if cell:
            row, col = cell
            try:
                elem.SetCellText(row, col, new_text)
            except Exception as e:
                output.print_md("❌ Failed to update cell ({}, {}) in `{}`: {}".format(row, col, elem.Name, e))
        else:
            try:
                elem.Text = new_text
            except Exception as e:
                output.print_md("❌ Failed to update TextNote: {}".format(e))
    t.Commit()
    forms.alert("✅ All changes saved.", title="Spell Checker")

def get_all_textnotes(doc):
    return FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TextNotes).WhereElementIsNotElementType().ToElements()

def get_all_schedules(doc):
    return FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()

def run_spell_checker():
    doc = __revit__.ActiveUIDocument.Document
    output.print_md("## ✏️ Spell Checker Started")

    for note in get_all_textnotes(doc):
        result = check_spelling(note.Text, "TextNote", element=note)
        if result == "STOP_COMMAND":
            break

    for sched in get_all_schedules(doc):
        output.print_md("📊 Checking Schedule: `{}`".format(sched.Name))
        try:
            table_data = sched.GetTableData()
            body_data = table_data.GetSectionData(1)  # 1 = Body
            n_rows = body_data.NumberOfRows
            n_cols = body_data.NumberOfColumns

            for row in range(n_rows):
                for col in range(n_cols):
                    try:
                        cell_text = sched.GetCellText(row, col)
                        result = check_spelling(cell_text, "Schedule: {}".format(sched.Name), element=sched, cell_pos=(row, col))
                        if result == "STOP_COMMAND":
                            raise StopIteration
                    except Exception as e:
                        output.print_md("⚠️ Skipped cell ({}, {}) in `{}`: {}".format(row, col, sched.Name, e))
        except StopIteration:
            break
        except Exception as e:
            output.print_md("⚠️ Could not process schedule `{}`: {}".format(sched.Name, e))

    if pending_edits:
        confirm = forms.alert("Do you want to save all changes made so far?", options=["Yes", "No"], exitscript=False)
        if confirm == "Yes":
            commit_changes()
        else:
            forms.alert("🚫 Changes discarded.", title="Spell Checker")
    else:
        forms.alert("No changes to save.", title="Spell Checker")

run_spell_checker()
