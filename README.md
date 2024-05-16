# SAP Business One 9.2 Item Attachment Console Application

## Summary

The following console application can be used to import into the SAP Business One Item Master Record one or more attachment records. This console application was written as this particular function is not available via the SAP DTW (Data Transfer Workbench) application.

This application uses a simple csv file as its data source. In essence, the console application reads each line in the file and attempts to add an attachment record and link this attachment to a item master record.

Your file can one or more attachments for a single Item Master record.

## Program Requirements

1. The code is based around Microsoft Visual Studio 2017 Community Edition. You will need this to run the code.
2. A SAP User Login to run the application. This user MUST have a professional license assigned.
3. You will need to include the SAPBobsCOM file in your project. This file is found in your SAP Business One Client Folder on your local machine.

![SAPBobsCom.png](https://bitbucket.org/repo/64My65M/images/741423303-SAPBobsCom.png)

## Before Running the Application

Before attempting to run this application the following changes are required.

1. A file Name and File Path. 

![OITM-Config-1.png](https://bitbucket.org/repo/64My65M/images/1078837868-OITM-Config-1.png)

2. SAP Login Details

![OITM-Config-2.png](https://bitbucket.org/repo/64My65M/images/462796140-OITM-Config-2.png)

## File Contents

1. The does not require a header record.
2. The file uses a single comma ',' as the delimiter. This however can be changed in the code to suit whatever you want to use. 

To change the delimiter see the GetFileContent Function


```
#!c#

var Fields = LineItem.Split(',');

```

File Contents is shown as below. A sample file is included in the source code directory called samplefile.csv


```
#!Text#

Item Code,File Path, File Name Without Extension, File Extension
---------------------------------------------------------------------------
0000001,\\10.61.1.15\share\SAP\Shared-Files\BUK\Images\www\HQ\,BK16,jpg

```


## Running the application

If you have installed or have installed visual studio 2017, you just run the application. It will walk through the code and upload item attachments and link these to the Item Master Record.

Here is a Before Attachment Sample Image

![OITM-Attach-1.png](https://bitbucket.org/repo/64My65M/images/3159955040-OITM-Attach-1.png)

Here is an After Attachment Sample Image

![OITM-Attach-2.png](https://bitbucket.org/repo/64My65M/images/1478807094-OITM-Attach-2.png)
