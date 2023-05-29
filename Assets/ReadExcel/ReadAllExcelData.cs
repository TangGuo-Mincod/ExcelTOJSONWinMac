using System.Collections.Generic;
using System.Data;
using System.IO;
using UnityEngine;
using Excel;
using LitJson;
using System.Reflection;
using System;
using System.Text.RegularExpressions;
using System.Text;
using UnityEditor;
using OfficeOpenXml;
using UnityEngine.TextCore;
using Unity.VisualScripting.Dependencies.NCalc;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.Xml.Serialization;

public class ReadAllExcelData : MonoBehaviour
{

    public static Action excelData;

    //类模板
    public static string classmoban = "using System;\r\nnamespace JsonAssets\r\n{\r\n    public class _ClassName \r\n    {\r\n      _body  \r\n    }\r\n}";
    //属性模板
    public static string classspromoban = "public _type _name;";

    /// <summary>
    /// 获取所有表
    /// </summary>
    public static void GetAllExcel()
    {

        DirectoryInfo info = new DirectoryInfo(MainExcel.exUrl);
        FileInfo[] excels = info.GetFiles("*");


        //遍历所有文件
        foreach (var item in excels)
        {
            if (!item.Name.EndsWith(MainExcel.exceltype))
            {
                continue;
            }
            //Debug.Log("文件名:" + item.Name);
            //Debug.Log("文件绝对路径:" + item.FullName);
            //Debug.Log("文件所在目录:" + item.DirectoryName);

            MainExcel.AddLog("读取excel 文件路径：" + item.FullName, Debuglog.finish);
            //获得表数据
            int columnNum = 0, rowNum = 0;
            //读取excel
            DataRowCollection collect = ReadExcel(item.FullName, ref columnNum, ref rowNum);

            MainExcel.AddLog("读取excel：" + "行  :" + rowNum + "列 :" + columnNum, Debuglog.finish);

            //////////////在内存中创建对象///////////////////
            //类名
            string scriptName = item.Name.Replace(MainExcel.exceltype, "");


            //创建对应类型的类文件
            CreatScriptFile(collect, columnNum, scriptName);



            //创建类的内容
            StringBuilder scriptBody = new StringBuilder();
            //遍历所有列
            for (int i = 0; i < columnNum; i++)
            {
                scriptBody.Append("\r\n");
                scriptBody.Append($"    " + 
                    classspromoban.Replace("_type", 
                    collect[2][i].ToString()).
                    Replace("_name", collect[1][i].ToString()));
                scriptBody.Append("\r\n");
                //类型
                Debug.Log("类型  :"+collect[0][i].ToString());
                //属性名
                Debug.Log("属性名 :"+collect[1][i].ToString());
            }

            //类型模板
            string cla = classmoban.Replace("_ClassName", scriptName).Replace("_body", scriptBody.ToString());


            MainExcel.AddLog("在内存中创建excel类型：" + cla, Debuglog.finish);
           
            List<object> JsonData = new List<object>();

           
            for (int i = 3; i < rowNum; i++)
            {
                object scriptsObj = Program.Creat(cla.Replace("number", "int"), scriptName);
                Type t = scriptsObj.GetType();
                for (int j = 0; j < columnNum; j++)
                {
                    //获取字段并赋值
                    FieldInfo scriptsfield = t.GetField(collect[1][j].ToString());
                    switch (collect[2][j].ToString())
                    {
                        case "int":
                            Debug.Log("int 类型的值  :"+ collect[i][j].ToString());
                            //赋值
                            scriptsfield.SetValue(scriptsObj,int.Parse(collect[i][j].ToString()));
                            break;
                        case "string":
                            //赋值
                            scriptsfield.SetValue(scriptsObj, collect[i][j].ToString());
                            break;
                        case "float":
                        case "double":
                            //赋值
                            scriptsfield.SetValue(scriptsObj, double.Parse(collect[i][j].ToString()));
                            break;
                        case "float[]":
                        case "double[]":
                            string[] data1 = collect[i][j].ToString().Split(",");
                            double[] intdata1 = new double[data1.Length];
                            for (int k = 0; k < data1.Length; k++)
                            {
                                intdata1[k] = double.Parse(data1[k]);
                            }
                            //赋值
                            scriptsfield.SetValue(scriptsObj, intdata1);
                            break;
                        case "int[]":
                            string[] data = collect[i][j].ToString().Split(",");
                            int[] intdata = new int[data.Length];
                            for (int k = 0; k < data.Length; k++)
                            {
                                intdata[k] = int.Parse(data[k]);
                            }
                            //赋值
                            scriptsfield.SetValue(scriptsObj, intdata);
                            break;
                        case "string[]":
                            //赋值
                            scriptsfield.SetValue(scriptsObj, collect[i][j].ToString().Split(","));
                            break;
                    }
                }
             
                JsonData.Add(scriptsObj);
            }

            DirectoryInfo di = new DirectoryInfo(MainExcel.jsonUrl + "/JsonFile/");
            if (!di.Exists)
            {
                di.Create();
            }
            string targetPath = MainExcel.jsonUrl + "/JsonFile/" + scriptName + ".json";
           // Debug.Log("创建文件的地址 :" + targetPath);
            FileInfo jsonFile = new FileInfo(targetPath);
            //如果文件存在删除重新创建
            if (jsonFile.Exists)
            {
                jsonFile.Delete();
                jsonFile.Create().Dispose();
            }
            else
            {
                jsonFile.Create().Dispose();
            }
            string jsonM = JsonMapper.ToJson(JsonData);
            Debug.Log(jsonM);
            //修改中文乱码问题
            Regex reg = new Regex(@"(?i)\\[uU]([0-9a-f]{4})");
            string jsonMDat = reg.Replace(jsonM, delegate (Match m) { return ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString(); });
            
            StreamWriter sw = new StreamWriter(targetPath);
            sw.Write(jsonMDat);
            sw.Dispose();
            sw.Close();
        }
        MainExcel.AddLog("<------- 导出完成 ------->" , Debuglog.finish);
        //打开文件夹
        ChinarMessage.ShellExecute(IntPtr.Zero, "open", MainExcel.jsonUrl, "", "", 1);
    }






    /// <summary>
    /// 读取excel文件内容
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="columnNum">行数</param>
    /// <param name="rowNum">列数</param>
    /// <returns></returns>
    private static DataRowCollection ReadExcel(string filePath, ref int columnNum, ref int rowNum)
    {
        FileStream stream = File.Open(filePath,FileMode.Open, FileAccess.Read, FileShare.Read);
        IExcelDataReader excelReader = null; //ExcelReaderFactory.CreateOpenXmlReader(stream);

  
        switch (MainExcel.exceltype)
        {
            case ".xls":
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                break;
            case ".xlsx":
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                break;
        }

        //返回的结果
        DataSet result = excelReader.AsDataSet();
        //Tables[0] 下标0表示excel文件中第一张表的数据
        columnNum = result.Tables[0].Columns.Count;
        rowNum = result.Tables[0].Rows.Count;
        return result.Tables[0].Rows;
    }


    /// <summary>
    /// 创建脚本文件
    /// </summary>
    private static void CreatScriptFile(DataRowCollection collect, int columnNum, string scriptName)
    {

        try
        {
            string moudle = MainExcel.GetMoudle();
            string[] all = moudle.Split('/');
            string cla = all[0];
            string shuxing = all[1];
            //创建类的内容
            StringBuilder scriptBody = new StringBuilder();
            //遍历所有列
            for (int i = 0; i < columnNum; i++)
            {
                //类型
                //Debug.Log("类型  :"+collect[0][i].ToString());
                //属性名
                //Debug.Log("属性名 :"+collect[1][i].ToString());
                scriptBody.Append("\r\n");

                scriptBody.Append($"    " + "/**\r\n     * " + collect[0][i].ToString() + " UI\r\n     */");

                if (MainExcel.type == "TypeScript")
                {
                    scriptBody.Append($"    " + shuxing.Replace("_type", collect[2][i].ToString().Replace("int", "number").Replace("double", "number")).Replace("_name", collect[1][i].ToString()));
                }
                else {
                    scriptBody.Append($"    " + shuxing.Replace("_type", collect[2][i].ToString()).Replace("_name", collect[1][i].ToString()));
                }
              
                scriptBody.Append("\r\n");
            }
            //类
            cla = cla.Replace("_ClassName", scriptName).Replace("_body", scriptBody.ToString());

            string scriptPath = MainExcel.jsonUrl + "/ScriptFile/";
            DirectoryInfo info = new DirectoryInfo(scriptPath);
            //如果文件夹不存在那么创建
            if (!info.Exists)
            {
                info.Create();
            }
            //判断文件是否存在 存在那么删除 重新创建
            if (File.Exists(scriptPath))
            {
                File.Delete(scriptPath);
                File.Create(scriptPath).Dispose();
            }
            else  //不存在直接创建
            {
                File.Create(scriptPath + scriptName + all[2]).Dispose();
            }
            File.WriteAllText(scriptPath + scriptName + all[2], cla);

            MainExcel.AddLog("写入数据完成 创建JSON文件 ：" + scriptName, Debuglog.finish);
        }
        catch (Exception)
        {
            throw;
        }

    }

}
