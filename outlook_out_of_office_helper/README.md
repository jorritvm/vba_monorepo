# Outlook Out of Office Scheduler

A small Outlook VBA macro that creates an out-of-office calendar block and daily vacation appointments in one click.

## What the macro does

When you run the macro, it:

1. Prompts for a start date and end date
2. Uses your selected calendar item or visible calendar range as a default if available
3. Creates an all-day Out of Office event covering the full date range
4. Creates one daily vacation appointment for each day between the selected dates

## How to add the macro to Outlook

1. In Outlook, click ALT+F11 to open the VBA Editor
2. In the VBA editor, right-click 'Modules' and choose 'Import File'
3. Select the `out_of_office.bas` file.
4. If Outlook asks about macro security, allow trusted macros or adjust the trusted locations in the Trust Center.

## Required personalization

Before running the macro, update the top-level constant in [out_of_office.bas](out_of_office.bas) so the out-of-office title matches your name:

`Private Const OOO_SUBJECT As String = "Your Name OoO"`

Replace `Your Name` with your own name.

## How to run it

- In Outlook, click ALT+F8
- Select the macro named CreateOoOAndVACAppointments.
- Click Run.

## How to add a shortcut to the Outlook Quick Access toolbar

1. In Outlook, click the dropdown next to the quick access toolbar and select 'More Commands'
2. On the left select 'Quick Access toolbar'  and then Choose commands from 'Macros'.
3. Find the 'CreateOoOAndVACAppointments' and move it to the right panel.
4. Click 'Modify' to select the desired Icon. 

This places a button directly in the Outlook Quick Access toolbar so you can start the macro with one click.
