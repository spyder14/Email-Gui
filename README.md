This is a powershell application to allow certain users to add and remove members from department email lists.
The user must already have permissions to the lists here "OU=365 Groups,OU=Groups,OU=PLP,OU=IFAS,OU=Departments,OU=UF,DC=ad,DC=ufl,DC=edu"

In order for the application to run, the computer needs to have some form of .Net framework to support the C# code used for sorting. 

It also needs Microsoft Remote Server Administration Tools(RSAT) so that it can leverage the powershell module "ActiveDirectroy" and its commandlettes.
This can be done by activating Windows Features, though WSUS settings prevent their installation unless worked around. Alternatively RSAT can be downloaded as a standalone app.

The computer needs to be able to execute unsigned powershell code. This can be done by running this command as admin "Set-ExecutionPolicy -ExecutionPolicy unrestricted"





This application was written by Michael Morrow
