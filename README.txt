
Powershell Oracle Word Generator
@Author: Felipe Donoso
         felipe@felipedonoso.cl
         felipe.donoso@oracle.com 
@Date: 2015-08-21

This is a very simple tool for create a word document with tables filled from oracle database querys. This is tool is built with powershell and is necessary have installed the oracle client (for the assembly System.Data.OracleClient)

The files necessary for generate the word are the next (Don't change the names:
- 00_FDB_Oracle_Word_Generator.cmd:  is for exec the report
- 01_FDB_Oracle_Word_Generator.ps1:  is the powershell code core and env variables (database string connection and name of every title or chapter)
- FDB_Oracle_Word_Generator__TEMPLATE03_NO_DELETE.docx: Template used for create the word base document

IMPORTANT: NO DELETE OR DROP STYLE INTO THE TEMPLATES FILES.

The execution of script it could be with these 2 command options:
1.- powershell -executionpolicy bypass -file 01_FDB_Oracle_Word_Generator.ps1
2.- executing this bat from CMD command line: 00_FDB_Oracle_Word_Generator.cmd

The output file it will be this name:
FDB_Oracle_Word_Generator_yyyymmddhh24miss_PDB_or_DBNAME.docx
