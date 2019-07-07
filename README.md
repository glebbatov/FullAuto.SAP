<p align="left">
<img src="https://github.com/glebbatov/FullAuto.SAP/blob/master/1200px-Microsoft_Office_Excel_(2018%E2%80%93present).svg.png" width="125">
  <h1>FullAuto.SAP</h1>
  <h3>[VBA, Excel, SAP ERP]</h3>
<p>
  
Automation solution based on Excel (VBA). Interacts with SAP ERP (enterprise resource planning) system for SAP client (production orders) at Insight (ten times decrease for production order process (hundred times stress reduction for a technician)).
#
<p>
<p align="left">
  <img src="https://github.com/glebbatov/FullAuto.SAP/blob/master/01.jpg" width="600">
  <img src="https://github.com/glebbatov/FullAuto.SAP/blob/master/02.jpg" width="600">
  <img src="https://github.com/glebbatov/FullAuto.SAP/blob/master/03.jpg" width="600">
</p>

# Versions
  
v.1.05

We here at SAP production care about our technicians’ sanity. We decided, that the most tedious process of working with SAP system must be replaced with a one button click.
Introducing the most up-to-date version of FullAuto™ v.1.05
* New FullAuto™ button "presses" first 4 buttons (Data Sheet):
	- Print LabNotes
	- Pull Order Quantity
	- Unpack
	-Print Stickers
* You can fix mistakes with new check boxes call "Reprint" in order to reprint LabNotes/Stickers second time (Data Sheet)
* Play Sound feature for FullAuto and CNFCheck buttons // requires *.wav file in the same folder as "FullAuto™ v.1.05.xlsm" (Data.J6 cell to change audio file name)
	- Control panel for play sound has been added (Data Sheet)
	- Check boxes ("Play") for play sound when FullAuto/CNFCheck clicked have been added (Data Sheet)

v.1.04
* Cell E12 allow choose a printer for stickers (Data)
* more cities/states to the location column have been added (Laptops)
* more devices have been added (Mobiles.UPC) 

v.1.03
* more cities/states to the location column have been added (Laptops)
* more devices have been added (Mobiles.UPC) 

v.1.02
* pulling data to userID/costCenter# columns, if the data is existing in SAP (Laptops/Mobile.Pull Data)
* script execution follows a current cell (Data.All buttons)
* wrong incrementing value has been fixed (Laptops/Mobile.Pull Data)

v.1.01
[first release version]
Production orders automation for SAP client
* Data sheet
	* Print LabNotes
	* Pull order quantity for pack/unpack
	* Pack/unpack orders
	* Print Stickers
	* Pull CNF Check (E18 cell changes time delay(in seconds))
* Laptops sheet
	* Pull data for orders from SAP system(sales order, used id, user name, email, shipping address, CostCenter#)
	* Variety of buttons for automation and smooth workflow
* Mobiles sheet
	* Pull data for orders from SAP system(sales order, used id, CostCenter#, user name, email, shipping address, area code, carrier, mobile number, PO# & Line Item)
	* Variety of buttons for automation and smooth workflow
* Additional sheets:
	* Laptops.MassDeployment (few clicks deployment)
	* Laptops.Location (filling Laptops.Location column automatically)
	* Mobile.UPC (filling Mobiles.DeviceType column automatically)
	* Stickers

# Developed By
Gleb Batov - batov.gleb1@gmail.com
