The PoolCleaner is a Microsoft Access database used to create a Microsoft Project Resource Pool.  A resource pool is a MS Project file that contains a list of resources that project managers can use in their schedules.  Project Managers link their schedules to the resource pool . This allows them to select resources from a unified file with a single format rather than typing resource names into their independent MS Project schedules.  By using a single resource pool across all project plans to select resources, analytics on project schedules related to resource information becomes significantly simpler because it eliminate the number of misspellings in people names as well as the opportunity to enter data entry errors.

The PoolCleaner takes an existing MS Project Resource Pool and connects to corporate informations systems that contain resource information, updates the information, and then create a new MS Project Resource Pool.  Information updates come from two corporate systems, Clarity an REI.  CA Clarity is a project and portfolio management system thetas often used for time tracking on projects.  REI is an employee information database that contains publicly available HR information that has been exporting from Lawson HR.  Using these two datasources employee and contractor name changes and status changes are automatically processed and added into the resource pool.  Typical status changes include adding new employees, adding new contractors, and making people as having left the company.  When people leave the company their records are kept in resource pool so the historical information in MS Project files isn’t lost.

PoolCleaner.mdb	- This is the MS Access database.

The following files are exports of the VBA scripts in the database.  Many of them contain embedded SQL that is runs to import and process data used in the processing the resource pool.

ErrHandler.vba			
Pool_Miscellaneous.vba
Form_Pool_frmResearch.vba
Form_frmGetInput.vba	
basImport.vba
Form_frmMain.vba	
basMain.vba			
modLocal_Relink_Fnctns_001.vba