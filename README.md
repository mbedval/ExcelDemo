#Setup the code

git clone https://github.com/mbedval/ExcelDemo.git
cd ExcelDemo
npm install
npm start

Note: npm start will start Excel installs locally, but it may not work depending on Restriction on Desktop Excel.

Testing/Demo the code on Office365 Excel.
--> Open the office 365 Excel
--> Select Menu [Insert --> Office Add-ins]
--> In Office-Add-ins Dialog Select Manage My Add-ins and upload the manifest.xml file from the project folder of ExcelDemo

Note: WebPack service should be running while demo which is run by command "npm start"
