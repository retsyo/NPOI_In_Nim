import sugar

import winim/clr

# in .NET Framework
var #DLL
    mscor = load("mscorlib")

var # Class
    String = mscor.GetType("System.String")
    FileStream = mscor.GetType("System.IO.FileStream")
    File = mscor.GetType("System.IO.File")

dump String
dump FileStream
dump File

# custom DLL
var #DLL
    NPOI = clr.load("NPOI.dll")
    NPOIOOXML = clr.load("NPOI.OOXML.dll")
    NPOIOpenXml4Net = clr.load("NPOI.OpenXml4Net.dll")
    NPOIOpenXmlFormats = clr.load("NPOI.OpenXmlFormats.dll")

dump NPOIOOXML

# Class
var XWPFDocument = NPOIOOXML.GetType("NPOI.XWPF.UserModel.XWPFDocument")
dump XWPFDocument

var doc = @XWPFDocument.new();

var table = doc.CreateTable(3, 3);

table.GetRow(0).GetCell(0).SetText("汉字测试");

var FileMode = mscor.GetType("System.IO.FileMode")
dump  @FileMode.Create[FileMode]

#~ var out1 = @FileStream.new(@String.new("table.docx"), @FileMode.Create[FileMode]);  # does not work
#~ var out1 = @FileStream.new("table.docx", 2);
var out1 = @FileStream.new("table.docx", @FileMode.Create[FileMode]); # this method works
#~ var out1 = @File.Create("table.docx");  # this method works

doc.Write(out1);
out1.Close();

