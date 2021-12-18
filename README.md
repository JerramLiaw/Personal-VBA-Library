# Personal VBA Library

This is a repository of all the VBA Macros that I have written either for myself or school. The purpose of this repository is to help me store an online version of the code as well as showcase what I have learnt in VBA so far.

As I am still a student, I expect that updating this Repo will be on a cyclical basis - I expect to update it more often during the Winter and Summer breaks as compared to the rest of the year.


## Version Control for VBA
Unfortunately, VBA has no in-built integration with [Git](https://git-scm.com/). Based on some research, there used to be a feature in the [Rubberduck add-in](https://rubberduckvba.com/) that incorporated VCS, but was [discontinued for some reason in 2019](https://stackoverflow.com/questions/41240745/version-control-system-for-excel-vba-code/41241438#41241438).

In order to have some form of version control for the rest of this repo, I started by building two macros - one to export my modules to a local Git repo and another to import them into the workbook from the repo. Since there is no way to merge or replace files, the macro clears off any existing modules in the repo before export and also deletes the modules from the workbook after export. I manually use Git in Powershell to commit changes onto the Local Repo and then push them onto Github. It is not a seamless method, but it works good enough for me.

This module was inspired by a [stackoverflow answer](https://stackoverflow.com/a/56630212). I adjusted the code to work better with Git and simplified several parts that would be redundant for my use case. Currently, the repo file path is hardcoded because of some issues with obtaining the file path if my desktop is synced to OneDrive. There are ways to [work around it](https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive), but I am still looking around if there are simpler methods.
