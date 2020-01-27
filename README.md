# vsto1
In short:
- the "add in" writes to a file on the local machine the selcted/ needed actions 
- a python aplication is tailing the "log file" and takes action based on some contract (did not create this)

How to run this POC: 
- install Visual Studio on a windows machine 
- Open this VSTO aplication 
- Click "Start"
- an "add in" will be install for your installed outlock aplication 
- open the "add in" and "click here "
- assumes this file "C:\Users\mldvn\test\WriteLines.txt" exits (you can hardcode any path here)
- will write 3 lines to the txt file

Resources:
- https://timdams.com/2017/05/09/how-to-create-a-simple-outlook-vsto-addin-a-step-by-step-guide/
- https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml
- https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Webapp/FunctionFile.html
- https://docs.microsoft.com/en-us/outlook/add-ins/add-in-commands-for-outlook
- https://docs.microsoft.com/en-us/outlook/add-ins/quick-start?tabs=visualstudio
- https://docs.microsoft.com/en-us/office/dev/add-ins/develop/develop-add-ins-visual-studio
- https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio?view=vs-2019
