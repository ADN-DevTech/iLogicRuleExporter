# iLogicRuleExporter
Utility to export all Rules from the Active document and all referenced documents to *.iLogicVb files

When you run it, it will find an active Inventor process, get the active document, and export all rules in it and in all referenced documents. The rules are exported as individual .iLogicVb files, in the same folder as the documents themselves. Generally it's a good idea to search for .iLogicVb files in the folders beforehand. But it's unlikely there will be name clashes because 
The files are generated with names in this format:
```
<AssemblyFilename>.iam.<RuleName>.iLogicVb
<PartFilename>.ipt.<RuleName>.iLogicVb
```

If a version of the file already exists and 
- it was generated by this program
- the current rule is not identical to the previous exported version
then the file will be overwritten.
Each file contains a comment line at the top that says it was exported by this program.

In the unlikely event that a file by this name already exists but it was *not* generated by this program, that file will not be overwritten. There is no warning in this case.
