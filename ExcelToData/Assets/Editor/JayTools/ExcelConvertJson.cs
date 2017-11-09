using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using UnityEditor;
using UnityEngine;

public class ExcelConvertJson : EditorWindow
{
    [MenuItem("Tools/ExcelToData %#~")]
    public static void ss()
    {
        string ExcelPath = Application.dataPath + "/Editor/JayTools/AssetImport.xlsx";
        string OutJsonFile = Application.dataPath + "/Scripts/TestJson.json";
        string OutCsFile = Application.dataPath + "/Scripts/TestClass.cs";
        WriteExcel(ExcelPath, OutJsonFile, OutCsFile, "TestClass");
    }


    public static void WriteExcel(string InPutFile, string OutJsonFile, string OutCsFile, string className)
    {
        //string outputDir = EditorUtility.SaveFilePanel("Save Excel", "", "New Resource", "xlsx");  
        FileInfo newFile = new FileInfo(InPutFile);
        if (!newFile.Exists)
            return;

        using (ExcelPackage package = new ExcelPackage(newFile))
        {
            ExcelWorksheet ImportPicSheet = GetExcelWorksheet("ImportPic", package);
            Dictionary<int, string> m_TitleDic = new Dictionary<int, string>();
            for (int i = 1; i < 1000; i++)
            {
                string value = ImportPicSheet.Cells[1, i].Text;

                if (string.IsNullOrEmpty(value))
                    break;

                if (value.StartsWith("*"))
                    continue;

                m_TitleDic.Add(i, value);
            }
            CreatJsonFile(OutJsonFile, ImportPicSheet, m_TitleDic);

            CreatCsFile(OutCsFile, ImportPicSheet, m_TitleDic, className);
        }

        AssetDatabase.SaveAssets();
        AssetDatabase.Refresh();
    }


    public const string StrCsFile = @"
[System.Serializable]
public class {0}
{{
{1}
}}";

    public static string StrData = "    public string {0};";

    private static string GetMainData(Dictionary<int, string> m_TitleDic)
    {
        StringBuilder sb = new StringBuilder();
        foreach (KeyValuePair<int, string> item in m_TitleDic)
        {
            string data = string.Format(StrData, item.Value);
            sb.AppendLine(data);
        }


        return sb.ToString();
    }

    private static void CreatCsFile(string OutFile, ExcelWorksheet ImportPicSheet, Dictionary<int, string> m_TitleDic, string FileName)
    {
        string MainData = GetMainData(m_TitleDic);
        string AllCsText = string.Format(StrCsFile, FileName, MainData);

        FileStream mFileStream = File.Create(OutFile);
        StreamWriter sw = new StreamWriter(mFileStream);

        sw.Write(AllCsText);
        sw.Close();
        sw.Dispose();
    }



    public static string MainStruct = 
@"//Excel To Json
//by Jay
[
{0}
]
";

    private static void CreatJsonFile(string OutFile, ExcelWorksheet ImportPicSheet, Dictionary<int, string> m_TitleDic)
    {
        FileStream mFileStream = File.Create(OutFile);
        StreamWriter sw = new StreamWriter(mFileStream);

        StringBuilder sb = new StringBuilder();
       
        //sb.AppendLine("[");

        for (int i = 2; i < 1000; i++)
        {
            string value = ImportPicSheet.Cells[i, 1].Text;
            if (string.IsNullOrEmpty(value))
                break;

            sb.AppendLine(GetLineJsonData(m_TitleDic, ImportPicSheet, i));
        }

        string jsonText =string.Format(MainStruct,sb.ToString());
        int LastIndex = jsonText.LastIndexOf(",");
        string FinalText = jsonText.Remove(LastIndex, 1);

        sw.Write(FinalText.ToString());
        sw.Close();
        sw.Dispose();
    }

    static string DataStruct = @"
  {{
{0}
  }},";

    static string LineData = "    \"{0}\": \"{1}\"";
    static string GetLineJsonData(Dictionary<int, string> m_TitleDic, ExcelWorksheet ImportPicSheet, int Line)
    {
        int ColumnTotalCount = m_TitleDic.Count;
        StringBuilder sb = new StringBuilder();

        foreach (KeyValuePair<int, string> item in m_TitleDic)
        {
            string Name = item.Value;
            string value = ImportPicSheet.Cells[Line, item.Key].Text;
            string lineStr = "";
            if (ColumnTotalCount == item.Key)//是否是最后一列的
            {
                lineStr = string.Format(LineData, Name, value);
            }
            else
            {
                lineStr = string.Format(LineData, Name, value) + ",";
            }
            sb.AppendLine(lineStr);
        }

        return string.Format(DataStruct, sb.ToString());
    }

    static ExcelWorksheet GetExcelWorksheet(string sheetName, ExcelPackage package)
    {
        foreach (ExcelWorksheet item in package.Workbook.Worksheets)
        {
            if (item.Name == sheetName)
                return item;
        }
        return null;
    }

}