 %UsageE:this .m file aim to copy several excel files to one excel in different
 %sheets ;make sure the first line is String and others are number in your
 %file or you should make changes 
 %CSDN blog:http://blog.csdn.net/thesunrize
 %Tips:you may well name your filename and path name in English in case
 %your matlab doesn't support Cinese
 %Coded by theSunrize ,2017.04.09
 %on matlabR2015b 
clear;
clc;

filePath=uigetdir({},'choose your filepath'); %get your file directory
getFileName=ls(strcat(filePath,'\*.xl*'));  %get the file name in your selected directory
fileName = cellstr(getFileName); %transfer string into cell array
 
if  isequal(getFileName,'')%make sure the directory you selected contains some excel files     
   msgbox('no excel file in the path you selected');
else 
 
mkdir(strcat(filePath,'\output'));%make a folder for output file
waiting=waitbar(0,'excuting...,please wait!');%make a waiting bar
for i=1: length(fileName)                                                  %foreach your files  
    [excelData,str] = xlsread(strcat(filePath,'\',fileName{i}));%get the string and data
    xlswrite(strcat(filePath,'\output\output.xlsx'),str,strcat('Sheet',num2str(i)), 'A1');%write title
    xlswrite(strcat(filePath,'\output\output.xlsx'),excelData,strcat('Sheet',num2str(i)), 'A2');%write data
   
end
close(waiting)%close waiting bar
disp 'you can find output file there:'
outputPath = strcat(filePath,'\output\output.xlsx')
msgbox(strcat('finished,get output file in:',outputPath),'Success','Help');%prompt message
end
 