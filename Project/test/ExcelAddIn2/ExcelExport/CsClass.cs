using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn2.ExcelExport
{
    public enum PublicType
    {
        PublicType_Public = 0,
        PublicType_Protected,
        PublicType_Private,
    }
   
    public class CsVariable
    {
        public class VariableType
        {
            public  string IntVar = "int";
            public  string floattVar = "float";
            public  string longVar = "long";
            public  string doubleVar = "double";
            public  string stringVar = "string"; 
        }

        public  PublicType m_PublicType = PublicType.PublicType_Public;
        public  string     m_Type;
        public  string     m_Name;
        public  string     m_Value;
        public string GetComplete()
        {
            string varString = "";
             string     PublicTypString = "";
            if(m_PublicType == PublicType.PublicType_Public)
            {
                PublicTypString = "public";
            }
            if(m_PublicType == PublicType.PublicType_Protected)
            {
                 PublicTypString = "protected";
            }
            if(m_PublicType == PublicType.PublicType_Private)
            {
                 PublicTypString = "private";
            }
            varString = PublicTypString;

            varString += m_Name + " ";
            if(string.IsNullOrEmpty(m_Value) == false)
            {
                 varString += " = " + m_Value + ";";
            }
            else
            {
                 varString +=";";
            }

            return varString;
        }
    }

    public class CsFunctionParam
    {
        public string m_Type;
        public string m_Name;
    }
    public enum CsFunctionOverrideType
    {
        None = 0,
        Override,
        Virtual,
        Interface,
    }

    public class CsFunctionContent
    {
        List<string> m_ContentLineList = new List<string>();
        public void Add(string contentLine)
        {
            m_ContentLineList.Add(contentLine);
        }
    }

    public class CsFunction
    {
        public PublicType m_PublicType = PublicType.PublicType_Public;
        public CsFunctionOverrideType m_CsFunctionOverrideType = CsFunctionOverrideType.None;
        public string m_ReturnType;
        public string m_Name;
        List<CsFunctionParam> m_ParamList = new List<CsFunctionParam>();
        public CsFunctionContent m_content;
        public void AddParam(CsFunctionParam param)
        {
            m_ParamList.Add(param);
        }
    }

    public class CsClass
    {
       public string ClassName;
       List<CsVariable> m_varList = new List<CsVariable>();
       List<CsFunction> m_varList = new List<CsFunction>();

       public void AddVar(CsVariable var)
       {
            
           m_varList.Add(var);
       }
    }
}
