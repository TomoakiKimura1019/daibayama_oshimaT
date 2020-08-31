using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KNdata
{
    class Class1
    {
        public static string strMid(string strValue, int intStartPos, int intCharCnt)
        {
            string strRet = string.Empty;
            try
            {
                strRet = strValue.Substring(intStartPos - 1, intCharCnt);
            }
            catch (Exception ex)
            {
                strRet = strValue;
            }
            return strRet;
        }


        public static string strRight(string strValue, int intLength)
        {
            string strRet = string.Empty;
            try
            {
                strRet = strValue.Substring(strValue.Length - intLength);
            }
            catch (Exception ex)
            {
                strRet = strValue;
            }
            return strRet;
        }

        public static string strLeft(string strValue, int intLength)
        {
            string strRet = string.Empty;
            try
            {
                if (string.IsNullOrEmpty(strValue)) strRet = strValue;
                intLength = Math.Abs(intLength);
                strRet = strValue.Length <= intLength ? strValue : strValue.Substring(0, intLength);
            }
            catch (Exception ex)
            {
                strRet = strValue;
            }
            return strRet;
        }  
    }

}
