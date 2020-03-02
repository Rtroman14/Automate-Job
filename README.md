# Automate Job

This project updates Excel database, and organizes specific data into a format for VFA (Asset Management Software) to read.

## Details

**MASTER TEMPLATE.xlsm** is the database we use to store all of our building systems and their respective attributes. After we collect data on assessed facilities, we list their systems and quantities into the "Systems" tab in "MASTER TEMPLATE.xlsm." Activating the "System Importer" button will execute **systemImporter.vs.**

**systemImporter.vs** creates a new sheet in **MASTER TEMPLATE.xlsm** which organizes each system with their respective quantity and matches up the correct attributes based on what is assigned in the "MASTER" tab. Once complete, we can upload this worksheet into VFA to populate our online database for our clients.

Two uploads to VFA are require to complete the task. **systemImporter.vs** creates the system in VFA, while **systemLineItems.vs** adds the correct costs to that system in VFA. Next, we have to export those systems from VFA to get a list of those systems with their unique ID. Once exported, activate "Line Item Importer" button on the "MASTER" tab in **MASTER TEMPLATE.xlsm** to execute its code. This will create another worksheet with each systems cost codes. Upload this to VFA to complete the task.
