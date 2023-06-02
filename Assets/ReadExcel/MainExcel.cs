using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

using UnityEngine;
using UnityEngine.UI;

public class MainExcel : MonoBehaviour
{


#if UNITY_EDITOR_OSX || UNITY_STANDALONE_OSX
    [DllImport("OpenFinderForUnity", CharSet = CharSet.Auto)]
    private static extern string getFilePath();

#endif

    public static string exUrl //表格的路径
    {
        get => PlayerPrefs.GetString("MainExcelexUrl", "");
        set => PlayerPrefs.SetString("MainExcelexUrl", value);
    }
    public static string jsonUrl  //Json文件和脚本文件的路径
    {
        get => PlayerPrefs.GetString("MainExceljsonUrl", "");
        set => PlayerPrefs.SetString("MainExceljsonUrl", value);
    }

    public static string moudleUrl  //类模板的路径
    {
        get => PlayerPrefs.GetString("MainExcelmoudleUrl", "");
        set => PlayerPrefs.SetString("MainExcelmoudleUrl", value);
    }

    public static string type   //脚本类型
    {
        get => PlayerPrefs.GetString("MainExceltype", "C#");
        set => PlayerPrefs.SetString("MainExceltype", value);
    }
    public static string exceltype  //表格类型
    {
        get => PlayerPrefs.GetString("MainExcelexceltype", ".xls");
        set => PlayerPrefs.SetString("MainExcelexceltype", value);
    }


    //本地模板
    public Dictionary<string, string> thislocalityMoudel;


    //选择excel路径
    public Button exurlButton;
    public InputField exurlInput;

    //输出路径
    public Button ExportButton;
    public InputField ExportInput;


    //脚本模板的路径
    public Button boardButton;//选择模板路径
    public Button refboardButton;//刷新模板
    public InputField boardInput;


    //选择脚本类型
    public Dropdown typeDro;

    //选择excel类型
    public Dropdown typeExcel;


    public Button DAOCHU;


    public Button exit;

    //日志: 
    public Transform content;


    //添加日志
    public static Action<string, Debuglog> AddLog; //添加日志

    //获取当前的模板
    public static Func<string> GetMoudle;


    private void Start()
    {
        
        GetMoudle = GetCurrMoudle;
        AddLog = AddDeBugLog;

        //读取模板
        ReadMoudle();
        if (exUrl != "")
        {
            exurlInput.text = exUrl;
        }

        if (jsonUrl != "")
        {
            ExportInput.text = jsonUrl;
        }

        if (moudleUrl!="")
        {
            boardInput.text = moudleUrl;
        }

        for (int i = 0; i < typeDro.options.Count; i++)
        {
            if (typeDro.options[i].text == type)
            {
                typeDro.value = i;
                break;
            }
        }

        for (int i = 0; i < typeExcel.options.Count; i++)
        {
            if (typeExcel.options[i].text == exceltype)
            {
                typeExcel.value = i;
                break;
            }
        }
        exit.onClick.AddListener(() => {
            PlayerPrefs.DeleteAll();
            Application.Quit();
        });

        DAOCHU.onClick.AddListener(() =>
        {
            ReadAllExcelData.GetAllExcel();
        });

        exurlButton.onClick.AddListener(() =>
        {

#if UNITY_EDITOR_OSX || UNITY_STANDALONE_OSX
           exUrl=getFilePath();
#else 
           exUrl = OpenFile.OpenWinFile();
#endif
            exurlInput.text = exUrl;
        });

        ExportButton.onClick.AddListener(() =>
        {
#if UNITY_EDITOR_OSX || UNITY_STANDALONE_OSX
            jsonUrl = getFilePath();
#else 
           jsonUrl = OpenFile.OpenWinFile();
#endif
            ExportInput.text = jsonUrl;
        });

        typeDro.onValueChanged.AddListener((id) =>
        {
            type = typeDro.options[id].text;
        });

        typeExcel.onValueChanged.AddListener((id) =>
        {
            exceltype = typeExcel.options[id].text;
        });


        //选择模板路径
        boardButton.onClick.AddListener(() =>
        {
            moudleUrl = OpenFile.OpenWinFile();
            boardInput.text = moudleUrl;
            ReadMoudle();
        });

        //路径输入的话需要用这个按钮刷新
        refboardButton.onClick.AddListener(() => {
            moudleUrl = boardInput.text;
            ReadMoudle();
        });
    }


    public void AddDeBugLog(string body, Debuglog log)
    {
        GameObject bodyText = Instantiate(Resources.Load<GameObject>("DebugLogText"));

        switch (log)
        {
            case Debuglog.err:
                body = "err : " + body;
                bodyText.GetComponent<Text>().color = Color.red;
                break;
            case Debuglog.finish:
                body = "finish : " + body;
                bodyText.GetComponent<Text>().color = Color.green;
                break;
        }
        bodyText.GetComponent<Text>().text = body;
        bodyText.transform.SetParent(content);
    }




    //返回当前选中的模板
    public string GetCurrMoudle() {
        return thislocalityMoudel[type];
    }

    public void ReadMoudle()
    {
        if (thislocalityMoudel == null)
        {
            thislocalityMoudel = new Dictionary<string, string>();
        }
        //加载本地模板
        TextAsset text = Resources.Load<TextAsset>("C#");
        if (!thislocalityMoudel.ContainsKey("C#"))
        {
            thislocalityMoudel.Add("C#", text.text);
            typeDro.AddOptions(new List<string>() { "C#" });
        }


        //加载本地模板
        TextAsset ts = Resources.Load<TextAsset>("TypeScript");
        if (!thislocalityMoudel.ContainsKey("TypeScript"))
        {
            thislocalityMoudel.Add("TypeScript", ts.text);
            typeDro.AddOptions(new List<string>() { "TypeScript" });

        }


        //开始加载模板
        if (moudleUrl != "")
        {
            DirectoryInfo info = new DirectoryInfo(moudleUrl);
            FileInfo[] excels = info.GetFiles("*");

            foreach (var item in excels)
            {
                if (!item.Name.EndsWith(".txt")|| thislocalityMoudel.ContainsKey(item.Name.Replace(".txt", "")))
                {
                    continue;
                }
                FileStream fs = item.OpenRead();

                StreamReader sr = new StreamReader(fs, System.Text.Encoding.Default);
                if (null == sr) return;
               
                string str = sr.ReadToEnd();
                thislocalityMoudel.Add(item.Name.Replace(".txt", ""), str);
 
                AddDeBugLog("外部类模板 -->  [" + item.Name + " :导入]", Debuglog.finish);
                fs.Close();
                sr.Close();
            }
            if (thislocalityMoudel == null|| thislocalityMoudel.Count<=0)
            {
                AddDeBugLog("<--------- 没有外部类模板 --------->", Debuglog.finish);
                return;
            }
            typeDro.AddOptions(new List<string>(thislocalityMoudel.Keys));
        }
        else
        {
            AddDeBugLog("<--------- 没有外部类模板 --------->", Debuglog.finish);
        }

    }


}


public enum Debuglog
{
    err,
    finish
}