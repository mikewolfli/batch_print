# batch_print
This project is for batch printing of label. It invokes the bartend dll. The data is snapped from web(used html parser).
Because the pythonnet can not use the dotnet 2.0 in win7(the bartend dll is dotnet 2.0), it can not compatible to dotnet 4.0. So I rebuild the pythonnet sample code with python 3.4 which is revised to be compatible to dotnet 2.0. and use this exe file to run the py file.(npython.exe)
It also used BeautifulSoup (but this can not run in npython.exe). so i change to the htmlparser.
but the beautifusoup sample is also kept in Beautifusoup folder.
