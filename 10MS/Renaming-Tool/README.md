# About Reanming-Tool

This is simple tool helps you to rename your files in a folder based on a sequence.

### Explanation
Say you have two folders: `folder-a` and `folder-b`.
Both of them have some files and their names follow the same sequence (ie. alphanumerically sortable)

```bash
$ tree folder-a
folder-a
├── ১)ত্রিকোণমিতি পর্ব  ১- (কোণ পরিমাপের পদ্ধতি).mp4
├── ১০)ত্রিকোণমিতি পর্ব  ১০.mp4
├── ১১)ত্রিকোণমিতি পর্ব  ১১.mp4
├── ২)ত্রিকোণমিতি পর্ব  ২- (রেডিয়ান কোণ,প্রতিজ্ঞা ১,২,৩,৪).mp4
├── ৭)ত্রিকোণমিতি পর্ব ৭-( গাণিতিক উদাহরণ ৮.৩ পর্ব ১).mp4
└── ৮)ত্রিকোণমিতি পর্ব ৮-( গাণিতিক উদাহরণ ৮.৩ পর্ব ২).mp4

$ tree folder-b
folder-b
├── file_1.mp4
├── file_10.mp4
├── file_11.mp4
├── file_2.mp4
├── file_7.mp4
└── file_8.mp4g
```

With help of `Renaming Tool` you can rename all the files of `folder-b` mapping to all the file names of `folder-a`.
In this case: `file_10` will get the name of `১০)ত্রিকোণমিতি পর্ব  ১০.mp4`, `file_2` will have the name of `২)ত্রিকোণমিতি পর্ব  ২- (রেডিয়ান কোণ,প্রতিজ্ঞা ১,২,৩,৪).mp4` and so on. Thus after hitting the `Rename` button in `RenamingTool` app the contents of folder-b will become:

```bash
$ tree folder-b
folder-b
├── report.csv
├── ১)ত্রিকোণমিতি পর্ব  ১- (কোণ পরিমাপের পদ্ধতি).mp4
├── ১০)ত্রিকোণমিতি পর্ব  ১০.mp4
├── ১১)ত্রিকোণমিতি পর্ব  ১১.mp4
├── ২)ত্রিকোণমিতি পর্ব  ২- (রেডিয়ান কোণ,প্রতিজ্ঞা ১,২,৩,৪).mp4
├── ৭)ত্রিকোণমিতি পর্ব ৭-( গাণিতিক উদাহরণ ৮.৩ পর্ব ১).mp4
└── ৮)ত্রিকোণমিতি পর্ব ৮-( গাণিতিক উদাহরণ ৮.৩ পর্ব ২).mp4

```
As you can see all the file names map to its name source folder `folder-a`. In the `RenamingTool` app it this folder is represented as `Input Folder` and `folder-b` ie the folder that uses the name of `folder-a` is termed as `Output Folder`.
Notice that there is also an extra file named `report.csv`, this file is autogenarated, will provide you which file name is mapped to which. If you don't want this file to be generated you can turn it of via the application gui.
If you want to delete `folder-a` after execution, you all also tick that in the gui.

### How to install?
Installation process is super easy.<br>
Just click on [`Setup.exe`](https://github.com/s-shifat/Automation-Scripts/blob/main/10MS/Renaming-Tool/Setup.exe) then click on the `Download` button.

Just make sure, you allow your browser to download `.exe` file.

##### Notes
* Currently the tool can sort English and Bengali only.

* Incase of Bengali file names. The file name should start with a `bangla number` followed by a `)` (similar to the tree above) then rest may conatain the file name.
* If you find any bug feel free to contact me at sshifat022@gmail.com

