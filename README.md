# Personal VBA Library

This is a repository of all the VBA Macros that I have written either for myself or school. The purpose of this repository is to help me store an online version of the code as well as **showcase what I have learnt in VBA so far**.

As I am still a student, I expect that updating this Repo will be on a cyclical basis - I expect to update it more often during the Winter and Summer breaks as compared to the rest of the year.


## Module 1: Version Control for VBA
Unfortunately, VBA has **no in-built integration with [Git](https://git-scm.com/)**. Based on some research, there used to be a feature in the [Rubberduck add-in](https://rubberduckvba.com/) that incorporated VCS, but was [discontinued for some reason in 2019](https://stackoverflow.com/questions/41240745/version-control-system-for-excel-vba-code/41241438#41241438).

In order to have some form of version control for the rest of this repo, I started by building two macros - one to **export my modules to a local Git repo and another to import them into the workbook from the repo**. Since there is no way to merge or replace files, the macro clears off any existing modules in the repo before export and also deletes the modules from the workbook after export. I then **manually use Git in Powershell** to commit changes onto the Local Repo and then push them onto Github. It is not a seamless method, but it works good enough for me.

This module was inspired by a [stackoverflow answer](https://stackoverflow.com/a/56630212). I adjusted the code with the intention of using Git and simplified several parts that would be redundant for my use case. Currently, the repo file path is hardcoded because of some issues with obtaining the file path if my desktop is synced to OneDrive. There are ways to [work around it](https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive), but I am still looking around if there are simpler methods.

Through coding this module, I gained **exposure to objects in the VB editor** in order to execute the import and export process. I also learnt how to deal with **File Paths and Scripting Objects** in order to know where and decide which modules to import or export.

## Module 2: Better Mail Merge
[Mail Merge](https://support.microsoft.com/en-us/office/use-mail-merge-for-bulk-email-letters-labels-and-envelopes-f488ed5b-b849-4c11-9cff-932c49474705) is a cross-platform feature that uses Microsoft Word, Outlook and Excel to **send personalized Emails in bulk**. However, there are some limitations - CC/BCC fields cannot be filled in, no attachments etc.

In order to overcome these limitations, I build a macro that **performs an improved mail merge as well as a macro that creates the template necessary for it to run**. The mail merge macro runs a for each loop that opens up Outlook and fills in the relevant sections from the details inside the Excel workbook as well as the body of the word document. Several parts of the email body are then changed using the **replace function** to reflect what is shown in the workbook, allowing for customized emails.

This module was inspired by an [old blogpost](http://exceltalk.blogspot.com/2014/03/customized-mail-merge-using-vba-in-word.html?m=1) that tried to do something similar. I adjusted the code with the intention of making the merge process fully customizable, not just stopping at CC/BCC. I also added an additional macro that would create a dynamic excel template whenever we want to run the macro. Similar to my version control module, my file paths are hardcoded due to limitations with using OneDrive.

Through coding this module, it helped me to **properly understand the Object Oriented Programming (OOP) way of coding and thinking** as this macro had to interact with different applications all at once. It also helped me to **understand loops** a lot better as they were extensively used to ensure the code was ran a precise number of times.
