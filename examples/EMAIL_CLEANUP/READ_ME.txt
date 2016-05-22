PURPOSE: quick cleanup of excel file for mailman mass subscription.

REQUIREMENTS: Windows machine (runs a VB script), Excel installed (2007 or greater), and Enable VBA access to macros in EXCEL trust center settings.

HOW TO USE: Place input xlsx file in same folder that "EmailListMaker.vbs" is in and then double click the"EmailListMaker" icon to run, this will produce (or recreate) an email_list.txt file that can be used for appending in mass subscription.


INPUT:

Expected input file to be an xlsx file (Excel) and have these columns (doesn't matter if header/names of columns change, but they need to be in this order):

First Name, Last Name, Preferred emails (up to 4)
-or, easier way to think of it-
First name, last name, email1, email2, email3, email4 


The name of the excel file doesn't matter (as long as it's xlsx).


Tested and planned for these types of names:
names containing tildes (~) and apostraphes(').
Commas throw everything off, so commas are removed from names.

The resulting file is email_list.txt and can be used to load up into the mailman list manager.

Takes preference of email columns in this order (first)C=>D=>E=>F (last - if no other emails are available), removes entries that have no emails.

OUTPUT:
Text file "email_list.txt" with one entry per line in the format:

Adam Smith <smithy@email.com>

