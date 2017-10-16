using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace ConsoleApp1
{

    class Program
    {
        class SignData
        {
            public string strDepartment;    //部门
            public string strName;          //姓名
            public DateTime aDateTime;      //时间
            public DateTime aDateDay;       //时间：天
            public string strDateDay;       //时间：天
            public string strDateHour;      //时间：小时
            public string strComment;       //备注
            public string strDescribe = "";      //描述

            public string FirstSignInTime = "无";//第一次签到时间
            public string FirstSignOutTime = "无";//第一次签退时间

            public string SecondSignInTime = "";//第二次签到时间
            public string SecondSignOutTime = "";//第二次签退时间

            public List<object> GetListObject()
            {
                List<object> lt = new List<object>();
                lt.Add(strDepartment);
                lt.Add(strName);
                lt.Add(strDateDay);
                lt.Add(strComment);
                lt.Add(FirstSignInTime);
                lt.Add(FirstSignOutTime);
                if (SecondSignInTime != "")
                {
                    lt.Add(SecondSignInTime);
                    lt.Add(SecondSignOutTime);
                }
                lt.Add(strDescribe);
                return lt;
            }
        }
        static void Main(string[] args)
        {
            //string aPath = "C:\\Users\\Sheng\\Desktop\\ConsoleApp1\\ConsoleApp1\\bin\\Debug\\";
            //Dictionary<object, List<List<object>>> aa = new Dictionary<object, List<List<object>>>();
            //SaveExcel(aPath, "new2.xlsx", ref aa);
            string aPath = System.IO.Directory.GetCurrentDirectory() + "\\";
            OpenExcel(aPath, "in");
        }

        static private void OpenExcel(string strFullPath, string strFileName)
        {
            object missing = System.Reflection.Missing.Value;
            Application excel = new Application();//lauch excel application
            if (excel == null)
            {
                Console.Write("<script>alert('Can't access excel')</script>");
            }
            else
            {
                Console.Write("读取数据中。。。\n");
                excel.Visible = false; excel.UserControl = true;
                // 以只读的形式打开EXCEL文件
                Workbook wb = excel.Application.Workbooks.Open(strFullPath + strFileName, missing, true, missing, missing, missing,
                 missing, missing, missing, true, missing, missing, missing, missing, missing);
                //取得第一个工作薄
                Worksheet ws = (Worksheet)wb.Worksheets.get_Item(1);


                //取得总记录行数   (包括标题列)
                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数
                int columnsint = ws.UsedRange.Cells.Columns.Count;//得到列数


                //取得数据范围区域 (不包括标题列) 
                Range rng1 = ws.Cells.get_Range("A2", "F" + rowsint);   //item

                //把Value2
                object[,] arryItem = (object[,])rng1.Value2;   //get range's value

                //将数组写入字典
                Dictionary<object, List<List<object>>> m_aTemp = new Dictionary<object, List<List<object>>>();
                for (int i = 1; i <= rowsint - 1; i++)
                {
                    object Okey = arryItem[i, 1];
                    if(!m_aTemp.ContainsKey(Okey))
                    {
                        m_aTemp.Add(Okey, new List<List<object>>());
                    }
                    List<object> ltObj = new List<object>();
                    for(int j = 1; j <= columnsint; j++)
                    {
                        ltObj.Add(arryItem[i, j]);
                    }
                    m_aTemp[Okey].Add(ltObj);

                }

                // 翻转一下列表，让日期小的在前面
                foreach (var strKey in m_aTemp.Keys)
                {
                    m_aTemp[strKey].Reverse();
                }

                Dictionary<object, List<List<object>>> m_aNew = new Dictionary<object, List<List<object>>>();
                
                foreach (var strKey in m_aTemp.Keys)
                {
                    if(!m_aNew.ContainsKey(strKey))
                    {
                        m_aNew.Add(strKey, new List<List<object>>());
                    }
                    List<SignData> ltSignData = new List<SignData>();

                    List<List<object>>lt = m_aTemp[strKey];
                    for (int i = 0; i < lt.Count; ++i)
                    {
                        ReadBase(i, ref lt, ref ltSignData);
                    }
                    List<SignData> ltDel = new List<SignData>();
                    for (int j = 0; j < ltSignData.Count -1 ; ++j)
                    {
                        SignData aSignData = ltSignData[j];
                        for (int k = j + 1; k < ltSignData.Count; ++k)
                        {
                            SignData _aSignData = ltSignData[k];
                            if (aSignData.aDateDay == _aSignData.aDateDay
                                && !aSignData.strComment.Contains("补签")
                                && !_aSignData.strComment.Contains("补签"))
                            {
                                aSignData.SecondSignInTime = _aSignData.FirstSignInTime;
                                aSignData.SecondSignOutTime = _aSignData.FirstSignOutTime;
                                aSignData.strDescribe += "  "+_aSignData.strDescribe;
                                ltDel.Add(_aSignData);
                            }
                        }
                        //m_aNew[strKey].Add(ltSignData[j].GetListObject());
                    }

                    foreach (var aDel in ltDel)
                    {
                        ltSignData.Remove(aDel);
                    }

                    foreach (var aData in ltSignData)
                    {
                        m_aNew[strKey].Add(aData.GetListObject());
                    }
                }


                excel.Quit();
                excel = null;
                Console.Write("读取数据完毕 \n");
                SaveExcel(strFullPath, "out.xlsx", ref m_aNew);
            }
        }

        static private SignData GetSignData(List<object> ltnew)
        {
            if (null == ltnew)
            {
                return null;
            }
            SignData aSignData = new SignData();

            string aL = ltnew[5].ToString();
            string strTime = ltnew[2].ToString();
            string strMin = ltnew[3].ToString();
            double dTime = double.Parse(strTime) + double.Parse(strMin);
            DateTime aDateTime = DateTime.FromOADate(dTime);
            DateTime aDay = aDateTime.Date;

            strTime = aDateTime.ToString();
            string[] arrTime = strTime.Split(' ');
            aSignData.strComment = aL;
            aSignData.aDateTime = aDateTime;
            aSignData.strName = ltnew[0].ToString();
            aSignData.strDepartment = ltnew[1].ToString();
            aSignData.aDateDay = aDay;
            aSignData.strDateDay = arrTime[0];
            aSignData.strDateHour = arrTime[1];

            return aSignData;
        }
        static private void ReadBase(int nIndex, ref List<List<object>> ltAll, ref List<SignData> ltSignData)
        {
            List<object> ltnew = ltAll[nIndex];
            SignData aSignData = GetSignData(ltnew);
            if (null == aSignData)
            {
                return;
            }

            
            if (aSignData.strComment.Contains("补签"))
            {
                aSignData.FirstSignInTime = aSignData.strDateHour;
                aSignData.strDescribe = "有补签，请检查";
                ltSignData.Add(aSignData);
            }
            else if(aSignData.strComment.Contains("签到"))
            {
                ReadSignIn(nIndex, ref ltAll, ref ltSignData, ref aSignData);
            }
            else if (aSignData.strComment.Contains("签退"))
            {
                ReadSignOut(nIndex, ref ltAll, ref ltSignData, ref aSignData);
            }
        }

        static private void ReadSignIn(int nIndex, ref List<List<object>> ltAll, ref List<SignData> ltSignData, ref SignData aSignData)
        {
            int NextIndex = nIndex + 1;
            if (NextIndex < ltAll.Count)
            {
                aSignData.FirstSignInTime = aSignData.strDateHour;
                SignData _aSignData = GetSignData(ltAll[NextIndex]);
                if (_aSignData.strComment.Contains("签退"))
                {
                    if (aSignData.aDateDay == _aSignData.aDateDay)
                    {
                        aSignData.FirstSignOutTime = _aSignData.strDateHour;
                        ltSignData.Add(aSignData);
                    }
                    else if (aSignData.aDateDay < _aSignData.aDateDay)
                    {
                        aSignData.FirstSignOutTime = "24:00";
                        ltSignData.Add(aSignData);

                        _aSignData.FirstSignInTime = "00:00";
                        _aSignData.FirstSignOutTime = _aSignData.strDateHour;
                        ltSignData.Add(_aSignData);
                    }
                }
                else
                {
                    aSignData.FirstSignOutTime = "无";
                    aSignData.strDescribe = "没有签退时间，请检查";
                    ltSignData.Add(aSignData);
                }
            }
        }

        static private void ReadSignOut(int nIndex, ref List<List<object>> ltAll, ref List<SignData> ltSignData, ref SignData aSignData)
        {
            foreach (var a in ltSignData)
            {
                if (a.aDateDay == aSignData.aDateDay)
                {
                    return;
                }
            }

            aSignData.FirstSignOutTime = aSignData.strDateHour;
            aSignData.strDescribe = "没有签到时间，请检查";
            ltSignData.Add(aSignData);
        }
        static private void SaveExcel(string strFullPath, string strFileName, ref Dictionary<object, List<List<object>>> map_Data)
        {
            Console.Write("保存数据中。。。。 \n");
            string strNewPath = strFullPath + strFileName;

            if (File.Exists(strNewPath))
            {
                File.Delete(strNewPath);
            }

            object Nothing = System.Reflection.Missing.Value;
            Application excel = new Application();//lauch excel application
            excel.Visible = false; excel.UserControl = true;
            Workbook wb = excel.Workbooks.Add(Nothing);
            Worksheet ws = wb.Sheets[1];
            int nCount = 0;
            foreach (var aValue in map_Data.Values)
            {
                foreach (var subValue in aValue)
                {
                    nCount++;
                    string strStart = "A" + nCount.ToString();
                    string strEnd = GetStrFormCount(subValue.Count) + nCount.ToString();
                    Range r = ws.get_Range(strStart, strEnd);
                    object[,] arryItem = (object[,])r.Value2;
                    for (int i = 1; i <= subValue.Count; ++i)
                    {
                        arryItem[1, i] = subValue[i - 1];
                    }
                    r.Value2 = arryItem;
                    
                }
            }
            
            object format = XlFileFormat.xlWorkbookDefault;
            Console.Write("写文件，时间比较长。耐心等待 \n");
            wb.SaveAs(strNewPath, format, Nothing, Nothing, Nothing, Nothing, XlSaveAsAccessMode.xlExclusive, Nothing, Nothing, Nothing, Nothing, Nothing);
            wb.Close(Nothing, Nothing, Nothing);
            excel.Quit();
            excel = null;
            Console.Write("保存数据完毕，按任意键退出 \n");
            Console.Read();
        }

        static string GetStrFormCount(int nCount)
        {
            switch(nCount)
            {
                case 1:
                    return "A";
                case 2:
                    return "B";
                case 3:
                    return "C";
                case 4:
                    return "D";
                case 5:
                    return "E";
                case 6:
                    return "F";
                case 7:
                    return "G";
                case 8:
                    return "H";
                case 9:
                    return "I";
                case 10:
                    return "J";
                case 11:
                    return "K";
                default:
                    return "A";
            }
        }
    }
}
