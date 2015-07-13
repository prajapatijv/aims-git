
select * from formcollection
--DROP TABLE formcollection

--------------MASTER FORMS----------------
Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMACMAST','ACCOUNT MASTER','ACCOUNT MASTER','M','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMCOMPMAST','COMPANY MASTER','COMPANY MASTER','M','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMCOMPOSITMAST','COMPOSITION MASTER','COMPOSITION MASTER','M','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMCOMPSEL','COMPANY SELECT','COMPANY SELECT','M','F')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMCOSTCENTMST','COST CENTER MASTER','COST CENTER MASTER','M','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMCOSTMETHODMAST','COST METHOD MASTER','COST METHOD MASTER','M','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMCRVIEWER','CRVIEWER','CRVIEWER','M','F')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMGDNMST','LOCATION MASTERS','LOCATION MASTERS','M','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMPALLETEMST','PLATE MASTER','PLATE MASTER','M','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMPATMAST','PATTERN MASTER','PATTERN MASTER','M','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMRAWMATMST','RAW MATERIAL MASTER','RAW MATERIAL MASTER','M','T')

--------------TRANSCTIONS FORMS----------------
Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMBATCHTRN','HEAT ENTRY','HEAT ENTRY','T','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMCHALTRN','CHALLAN ENTRY','CHALLAN ENTRY','T','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMORDTRN','SALES ORDER ENTRY','SALES ORDER ENTRY','T','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMPATDESPATCH','PATTERN DESPATCH ENTRY','PATTERN DESPATCH ENTRY','T','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMPATRECEIVE','PATTERN RECEIVE ENTRY','PATTERN RECEIVE ENTRY','T','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMPRODTRN','DAILY PRODUCTION ENTRY','DAILY PRODUCTION ENTRY','T','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMPURTRN','MATERIAL INWARD ENTRY','MATERIAL INWARD ENTRY','T','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMREJECTRN','REJECTION ENTRY','REJECTION ENTRY','T','T')


--------------REPORTS FORMS----------------
Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMREPSPROD','PRODUCTION & CONSUMPTION REPORTS','PRODUCTION & CONSUMPTION REPORTS','R','T')

Insert into FormCollection (FormName,FormCap,Descr,FormGrp,Status)
values ('FRMREPSPTRN','PATTERN REPORTS','PATTERN REPORTS','R','T')
