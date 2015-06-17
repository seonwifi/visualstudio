using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn2.ExcelExport
{

   
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

        public  PublicType m_PublicType = PublicType.PublicType_public;
        public  string     m_Type;
        public  string     m_Name;
        public  string     m_Value;

        public CsVariable(PublicType v_PublicType, string v_Type, string v_Name, string v_Value)
        {
            m_PublicType = v_PublicType;
            m_Type =v_Type;
            m_Name = v_Name;
            m_Value = v_Value;
        }

        public string MakeString()
        {
            string varString = "";
            string PublicTypString = CSEnumConvert.EnumToString<PublicType>(m_PublicType); 
            varString = PublicTypString;
            varString += " " + m_Type;
            varString += " " + m_Name;
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

        public CsFunctionParam(string v_Type, string v_Name)
        {
            m_Type = v_Type;
            m_Name = v_Name;
        }

        public string MakeString()
        {
            return m_Type + " " + m_Name;
        }
    }


   

    public class CsClass
    {
        static public int StringDepth = 0;
       public PublicType m_PublicType = PublicType.PublicType_public;
       public string m_ClassName;
       public string m_ParentClassName = "";
       List<CsVariable> m_varList = new List<CsVariable>();
       List<CsFunction> m_FunctionList = new List<CsFunction>();

       public CsClass(PublicType publicType, string className, string parentClassName = "")
       {
           m_PublicType = publicType;
           m_ClassName = className;
           m_ParentClassName = parentClassName;
       }

       public void AddVar(CsVariable var)
       {
            
           m_varList.Add(var);
       }

       public string DepthToString(int depth)
       {
           string varString = "";
           for (int i = 0; i < depth; ++i)
           {
               varString += "\t";
           }
           return varString;
       }

       public string MakeString()
       {
           StringDepth = 0;

           string varString = "";
           string PublicTypString = CSEnumConvert.EnumToString<PublicType>(m_PublicType);
           varString = PublicTypString;

           varString += " class " + m_ClassName;
            if(string.IsNullOrEmpty(m_ParentClassName) == false)
            {
                varString += " : public " + m_ParentClassName;
            }
            varString += "\r\n{";
 
           for (int i = 0; i < m_varList.Count; ++i)
           {
              varString += "\r\n";
              varString += DepthToString(StringDepth + 1) + m_varList[i].MakeString(); 
           }

           for (int i = 0; i < m_FunctionList.Count; ++i)
           {
               varString += "\r\n";
               varString += DepthToString(StringDepth + 1) + m_FunctionList[i].MakeString();
           }

           varString += "\r\n}";
           return varString;
       }
    }
}
