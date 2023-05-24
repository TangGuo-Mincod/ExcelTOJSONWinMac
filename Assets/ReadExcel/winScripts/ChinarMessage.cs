using System;
using System.Runtime.InteropServices;//调用外部库，需要引用命名空间
using UnityEngine;
using Debug = UnityEngine.Debug;

/// <summary>
/// 为了调用外部库脚本
/// </summary>
public static class ChinarMessage
{
    [DllImport("User32.dll", SetLastError = true, ThrowOnUnmappableChar = true, CharSet = CharSet.Auto)]
    public static extern int MessageBox(IntPtr handle, String message, String title, int type);//具体方法


    public static int returnNumber;

    /// <summary>
    /// 9个按钮对应弹框
    /// </summary>
    /// <param name="index"></param>
    private static void Button(int index)
    {
        switch (index)
        {
            case 0:
                returnNumber = ChinarMessage.MessageBox(IntPtr.Zero, "Chinar-0:返回值均：1", "确认", 0);
     
                break;
            case 1:
                returnNumber = ChinarMessage.MessageBox(IntPtr.Zero, "Chinar-1:确认：1，取消：2", "确认|取消", 1);
                break;
            case 2:
                returnNumber = ChinarMessage.MessageBox(IntPtr.Zero, "Chinar-2:中止：3，重试：4，忽略：5", "中止|重试|忽略", 2);
                break;
            case 3:
                returnNumber = ChinarMessage.MessageBox(IntPtr.Zero, "Chinar-3:是：6，否：7，取消：2", "是 | 否 | 取消", 3);
                break;
            case 4:
                returnNumber = ChinarMessage.MessageBox(IntPtr.Zero, "Chinar-4:是：6，否：7", "是 | 否", 4);
                break;
            case 5:
                returnNumber = ChinarMessage.MessageBox(IntPtr.Zero, "Chinar-5:重试：4，取消：2", "重试 | 取消", 5);
                break;
            case 6:
                returnNumber = ChinarMessage.MessageBox(IntPtr.Zero, "Chinar-6:取消：2，重试：10，继续：11", "取消 | 重试 | 继续", 6);
                break;
        }
    }


    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern int ShellExecute(IntPtr hwnd, string lpszOp, string lpszFile,string lpszParams, string lpszDir, int FsShowCmd);


}


public class OpenFile : MonoBehaviour
{
    public static string OpenWinFile()
    {
        //使用如下：
        OpenDialogDir ofn2 = new OpenDialogDir();
        ofn2.pszDisplayName = new string(new char[2048]); ;     // 存放目录路径缓冲区  
        ofn2.lpszTitle = "选择保存路径";// 标题  
        //ofn2.ulFlags = BIF_NEWDIALOGSTYLE | BIF_EDITBOX; // 新的样式,带编辑框  
        //IntPtr pidlPtr = IntPtr.Zero;
        IntPtr pidlPtr = DllOpenFileDialog.SHBrowseForFolder(ofn2);

        char[] charArray = new char[2048];
        for (int i = 0; i < 2048; i++)
            charArray[i] = '\0';

        DllOpenFileDialog.SHGetPathFromIDList(pidlPtr, charArray);
        string fullDirPath = new String(charArray);
        fullDirPath = fullDirPath.Substring(0, fullDirPath.IndexOf('\0'));

        Debug.Log(fullDirPath);//这个就是选择的目录路径
        return fullDirPath;
    }
  
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
public class OpenDialogFile
{
    public int structSize = 0;
    public IntPtr dlgOwner = IntPtr.Zero;
    public IntPtr instance = IntPtr.Zero;
    public String filter = null;
    public String customFilter = null;
    public int maxCustFilter = 0;
    public int filterIndex = 0;
    public String file = null;
    public int maxFile = 0;
    public String fileTitle = null;
    public int maxFileTitle = 0;
    public String initialDir = null;
    public String title = null;
    public int flags = 0;
    public short fileOffset = 0;
    public short fileExtension = 0;
    public String defExt = null;
    public IntPtr custData = IntPtr.Zero;
    public IntPtr hook = IntPtr.Zero;
    public String templateName = null;
    public IntPtr reservedPtr = IntPtr.Zero;
    public int reservedInt = 0;
    public int flagsEx = 0;
}
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
public class OpenDialogDir
{
    public IntPtr hwndOwner = IntPtr.Zero;
    public IntPtr pidlRoot = IntPtr.Zero;
    public String pszDisplayName = null;
    public String lpszTitle = null;
    public UInt32 ulFlags = 0;
    public IntPtr lpfn = IntPtr.Zero;
    public IntPtr lParam = IntPtr.Zero;
    public int iImage = 0;
}

public class DllOpenFileDialog
{
    [DllImport("Comdlg32.dll", SetLastError = true, ThrowOnUnmappableChar = true, CharSet = CharSet.Auto)]
    public static extern bool GetOpenFileName([In, Out] OpenDialogFile ofn);

    [DllImport("Comdlg32.dll", SetLastError = true, ThrowOnUnmappableChar = true, CharSet = CharSet.Auto)]
    public static extern bool GetSaveFileName([In, Out] OpenDialogFile ofn);

    [DllImport("shell32.dll", SetLastError = true, ThrowOnUnmappableChar = true, CharSet = CharSet.Auto)]
    public static extern IntPtr SHBrowseForFolder([In, Out] OpenDialogDir ofn);

    [DllImport("shell32.dll", SetLastError = true, ThrowOnUnmappableChar = true, CharSet = CharSet.Auto)]
    public static extern bool SHGetPathFromIDList([In] IntPtr pidl, [In, Out] char[] fileName);

}





