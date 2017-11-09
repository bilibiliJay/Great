
#if UNITY_EDITOR
using OfficeOpenXml;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;

public class JayExcel : MonoBehaviour
{

    void SS()
    {
        WriteExcel("");
    }

    public static void WriteExcel(string outputDir)
    {
        //string outputDir = EditorUtility.SaveFilePanel("Save Excel", "", "New Resource", "xlsx");  
        FileInfo newFile = new FileInfo(outputDir);
        if (newFile.Exists)
        {
            newFile.Delete();  // ensures we create a new workbook  
            newFile = new FileInfo(outputDir);
        }

        //List<AnimFileData> mList = BattleAnimation.GetAnimType(BattleAnimation.AnimPath);

        using (ExcelPackage package = new ExcelPackage(newFile))
        {
            // add a new worksheet to the empty workbook  
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
            //Add the headers  
            worksheet.Cells[1, 1].Value = "分类";
            worksheet.Cells[1, 2].Value = "前缀";
            worksheet.Cells[1, 3].Value = "下限";
            worksheet.Cells[1, 4].Value = "上限";

            worksheet.Cells[2, 1].Value = "Type";
            worksheet.Cells[2, 2].Value = "Prefix";
            worksheet.Cells[2, 3].Value = "Lower";
            worksheet.Cells[2, 4].Value = "Upper";

            worksheet.Cells[3, 1].Value = "int";
            worksheet.Cells[3, 2].Value = "string";
            worksheet.Cells[3, 3].Value = "int";
            worksheet.Cells[3, 4].Value = "int";

            int index = 4;
            //for (int i = 0; i < mList.Count; i++)
            {
                //AnimFileData mAnimationExcelData = mList[i];
              
                //worksheet.Cells[index, 1].Value = i+1;
                //worksheet.Cells[index, 2].Value = mAnimationExcelData.animType;
                worksheet.Cells[index, 3].Value = 1;
                //worksheet.Cells[index, 4].Value = mAnimationExcelData.mClips.Count;
                index++;
            }

            //Add some items...  
            //worksheet.Cells["A2"].Value = 12001;
            //save our new workbook and we are done!  
            package.Save();
        }
    }

    public static void WriteAnimationDataExcel(string outputDir)
    {
        //string outputDir = EditorUtility.SaveFilePanel("Save Excel", "", "New Resource", "xlsx");  
        FileInfo newFile = new FileInfo(outputDir);
        if (newFile.Exists)
        {
            newFile.Delete();  // ensures we create a new workbook  
            newFile = new FileInfo(outputDir);
        }
        using (ExcelPackage package = new ExcelPackage(newFile))
        {
            // add a new worksheet to the empty workbook  
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
            string tmp = worksheet.Cells[1, 1].Text;
            //Add the headers  
            worksheet.Cells[1, 1].Value = "分类";
            worksheet.Cells[1, 2].Value = "时长";
            //worksheet.Cells[1, 3].Value = "下限";
            //worksheet.Cells[1, 4].Value = "上限";

            worksheet.Cells[2, 1].Value = "Index";
            worksheet.Cells[2, 2].Value = "Time";
            //worksheet.Cells[2, 3].Value = "Lower";
            //worksheet.Cells[2, 4].Value = "Upper";

            worksheet.Cells[3, 1].Value = "string";
            worksheet.Cells[3, 2].Value = "double";
            //worksheet.Cells[3, 3].Value = "int";
            //worksheet.Cells[3, 4].Value = "int";

            //List<AnimFileData> mList = BattleAnimation.GetAnimType(BattleAnimation.AnimPath);
            int index = 4;
            //for (int i = 0; i < mList.Count; i++)
            {
                //AnimFileData mAnimationExcelData = mList[i];

                //foreach (AnimationClip mClip in mAnimationExcelData.mClips)
                //{
                //    worksheet.Cells[index, 1].Value = mClip.name;
                //    worksheet.Cells[index, 2].Value = System.Math.Round(mClip.length, 3);
                //    index++;
                //}
            }
            package.Save();
        }

        //Add some items...  
        //worksheet.Cells["A2"].Value = 12001;
        //save our new workbook and we are done!  
    }
}



public class AnimationExcelData
{
    public int m_Type;
    public string Prefix;
    public int Lower;
    public int Upper;

    public AnimationExcelData()
    {
        m_Type = Random.Range(1, 10);
        Prefix = "123123";
        Lower = Random.Range(1, 10);
        Upper = Random.Range(1, 10);
    }
}
#endif