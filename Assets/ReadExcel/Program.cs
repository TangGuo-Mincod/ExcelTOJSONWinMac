using System;
using System.Reflection;
using System.Globalization;
using Microsoft.CSharp;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Text;


public class Program
{
    //在内存中编译程序集 动态创对象
    public static object Creat(string cla,string name)
    {
        // 1.CSharpCodePrivoder 
        CSharpCodeProvider objCSharpCodePrivoder = new CSharpCodeProvider();
    
        // 2.ICodeComplier 
        ICodeCompiler objICodeCompiler = objCSharpCodePrivoder.CreateCompiler();

        // 3.CompilerParameters 
        CompilerParameters objCompilerParameters = new CompilerParameters();
        objCompilerParameters.ReferencedAssemblies.Add("System.dll");
        objCompilerParameters.GenerateExecutable = false;
        objCompilerParameters.GenerateInMemory = true;

        // 4.CompilerResults 
        CompilerResults cr = objICodeCompiler.CompileAssemblyFromSource(objCompilerParameters, cla);
   
        object objHelloWorld = null;
        if (cr.Errors.HasErrors)
        {
            foreach (CompilerError err in cr.Errors)
            {
                MainExcel.AddLog("内存创建程序集编译错误：" + err.ErrorText,Debuglog.err);
            }
        }
        else
        {
            // 通过反射，创建实例 
            Assembly objAssembly = cr.CompiledAssembly;
            //在新的命名空间中添加
            objHelloWorld = objAssembly.CreateInstance("JsonAssets."+name);
            MainExcel.AddLog("内存创建程序集成功 ：" + "JsonAssets." + name, Debuglog.finish);
        }
        return objHelloWorld;
    }
}
