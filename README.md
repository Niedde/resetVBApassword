# resetVBApassword
A piece of R code that can be useful if you forgot the password to your VBA project.

Including a protected Excel file to test the program.

*!! Don't use it to access files you're not supposed to !!*


### How to use it :

Run the first part of the code. R opens a window to let you choose the .xlsm file and makes a copy of the file which is saved as a .zip archive.
The .zip archive opens automatically. 

Open xl subfolder and find vbaProject.bin manually, then copy vbaProject.bin into another folder manually.

The second part of the code is here to edit the vbaProject.bin file by changing a key called DPB. Run it.

Put manually the modified vbaProject.bin file into the zip folder, then manually rename the .zip file to .xlsm and open it. Excel doesn't recognize the security anymore so you get a bunch of error messages "Invalid key", "System error 40230" .. Don't panic ! Just click "OK" as many times as needed.
Open VBA editor, then go to Tools > VBA Project Properties > Protection and enter a new password.

Save your Excel file, close it and re-open it. Voilà !
