using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.ComponentModel;
using System.Reflection;
using System.Text;

namespace EditablePivot.BaseClasses
{
    public class ListMethods
    {
        public static List<DataType> sortAndRemoveDuplicates<DataType>(List<DataType> lst)
        {
            List<DataType> ret = new List<DataType>();
            lst.Sort();
            for (int i = 0; i <= lst.Count - 1; i++)
            {
                if (!ret.Contains(lst[i]))
                {
                    ret.Add(lst[i]);
                }
            }
            return ret;
        }

        public static string ToString<DataType>(List<DataType> lst, string Delimiter = ", ")
        {
            string ret = "";
            for (int i = 0; i <= lst.Count - 1; i++)
            {
                if (ret != "") ret += Delimiter;
                ret += lst[i].ToString();
            }
            return ret;
        }

        public static bool CompareList(List<int> input1, List<int> input2)
        {
            List<int> lst1 = sortAndRemoveDuplicates<int>(input1);
            List<int> lst2 = sortAndRemoveDuplicates<int>(input2);

            if (lst1.Count != lst2.Count)
            {
                return false;
            }
            else
            {
                for (int i = 0; i <= lst1.Count - 1; i++)
                {
                    if (lst1[i] != lst2[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }
    }

    public class IdList : List<int>
    {
        public IdList() { }

        // initialize a IdList from a string with : delimited
        public IdList(string Ids)
        {
            if (Ids != null && Ids != "")
            {
                foreach (string Id in Ids.Split(GeneralSettings.IdDelimiterApp.ToCharArray()))
                {
                    if (Id.Trim() != "") this.Add(Convert.ToInt32(Id));
                }
            }
        }

        public int GetElement(int Id)
        {
            return this.Find(e => e == Id);
        }

        public string ToString(string delimiter = GeneralSettings.IdDelimiterApp)
        {
            string ret = "";
            foreach (int i in this)
            {
                if (ret != "") ret += delimiter;
                ret += i.ToString();
            }

            return ret;
        }
    }

    public class StringList : List<String>
    {
        public StringList() { }

        // initialize a IdList from a string with : delimited
        public StringList(string lst)
        {
            if (lst != null && lst != "")
            {
                this.AddRange(lst.Split(GeneralSettings.IdDelimiterApp.ToCharArray()));
            }
        }

        public string GetElement(String Id)
        {
            return this.Find(e => e == Id);
        }

        public string Print()
        {
            StringBuilder ret = new StringBuilder();
            foreach (string str in this)
            {
                if (ret.Length > 0) ret.Append(Environment.NewLine);
                ret.Append(str);
            }
            return ret.ToString();
        }

        public string ToString(string delimiter = GeneralSettings.IdDelimiterApp)
        {
            StringBuilder ret = new StringBuilder();
            foreach (string str in this)
            {
                if (ret.Length > 0) ret.Append(delimiter);
                ret.Append(str);
            }
            return ret.ToString();
        }
    }

    public class GeneralSettings
    {
        public static CultureInfo DefCulture = CultureInfo.CreateSpecificCulture("en-US");
        public const string DefDateTimeFormat = "yyyy-MM-dd HH:mm:ss";
        public const string DateOnlyFormat = "yyyy-MM-dd";

        public const string TimeOnlyFormat = "HH:mm:ss";
        public const string IdDelimiterApp = ":";

        public const string IdDelimiterDb = ",";
        public static DateTime MinDate = DateTime.ParseExact("1900-01-01", DateOnlyFormat, CultureInfo.InvariantCulture);
        public static DateTime MinDateTime = DateTime.ParseExact("1900-01-01 00:00:00", DefDateTimeFormat, CultureInfo.InvariantCulture);
        public static DateTime MaxDate = DateTime.ParseExact("2900-01-01", DateOnlyFormat, CultureInfo.InvariantCulture);
        // Date variable cannot be NULL, thus use 1900-01-01 as NULL

        public const string NullDateString = "1900-01-01";

        // constant charArray used for trim the line end
        public static char[] trimChars = new char[] { '\r', '\n' };

        public static bool IsNumeric(object expression)
        {
            if (expression == null)
                return false;
            if (expression.GetType() == typeof(string) && expression.ToString().Contains(CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator))
                // The C# Double.TryParse method handles group separators different than Excel. 
                // We need to handle them here to avoid counter-intuitive results.
                //    1,770 -> double.tryParser returns true :  OK
                //    1,77  -> double.tryParser returns false:  NOT OK, Excel interpretes it as an invalid number.
                return false;
            double number;
            return Double.TryParse(Convert.ToString(expression, CultureInfo.InvariantCulture), System.Globalization.NumberStyles.Any, NumberFormatInfo.InvariantInfo, out number);
        }

        public static bool IsDate(object expression)
        {
            if (expression == null) return false;

            string strDate = expression.ToString();
            try
            {
                DateTime dt = DateTime.Parse(strDate);
                if (dt != DateTime.MinValue && dt != DateTime.MaxValue)
                    return true;
                return false;
            }
            catch
            {
                return false;
            }
        }

        public static DateTime StringToDate(string str)
        {
            try
            {
                return DateTime.ParseExact(str, DateOnlyFormat, DefCulture);
            }
            catch (Exception ex)
            {
                throw new Exception("Fail to convert to Date: \n" + ex.Message);
            }
        }

        public static DateTime StringToDateTime(string str)
        {
            try
            {
                return DateTime.ParseExact(str, DefDateTimeFormat, DefCulture);
            }
            catch (Exception ex)
            {
                throw new Exception("Fail to convert to Date: \n" + ex.Message);
            }
        }

        public static DateTime CDate2(object obj)
        {
            if (object.ReferenceEquals(obj, DBNull.Value))
            {
                return DateTime.Parse(NullDateString);
            }
            else
            {
                return (DateTime)obj;
            }
        }

        public static string DateToString(DateTime? myDate, string fmt)
        {
            if (myDate == null) return "";
            if (myDate == DateTime.Parse(NullDateString))
            {
                return "";
            }
            else
            {
                return Convert.ToDateTime(myDate).ToString(fmt);
            }
        }

        public static string DateToString(DateTime? myDate)
        {
            if (myDate == null) return "";
            else return Convert.ToDateTime(myDate).ToString(GeneralSettings.DefDateTimeFormat);
        }

        public static DateTime GetRelativeDate(DateTime MyDate, string DateUnit, int RelativeValue, string FirstOrLast)
        {
            DateTime ret = default(DateTime);
            switch (DateUnit)
            {
                case "Day":
                    ret = MyDate.AddDays(RelativeValue);
                    break;
                case "Week":
                    ret = MyDate.AddDays((double)(1 - MyDate.DayOfWeek)).AddDays(7 * RelativeValue);
                    if (FirstOrLast == "Last")
                    {
                        ret = ret.AddDays(6);
                    }
                    break;
                case "Month":
                    ret = MyDate.AddDays(-MyDate.Day + 1).AddMonths(RelativeValue);
                    if (FirstOrLast == "Last")
                    {
                        ret = ret.AddMonths(1).AddDays(-1);
                    }
                    break;
                case "Quarter":

                    int quar = (int)Math.Floor((double)((MyDate.Month - 1) / 3)) + 1;
                    ret = new DateTime(MyDate.Year, 3 * quar - 2, 1).AddMonths(RelativeValue * 3);
                    if (FirstOrLast == "Last")
                    {
                        ret = ret.AddMonths(3).AddDays(-1);
                    }
                    break;
                case "Year":
                    ret = new DateTime(MyDate.Year, 1, 1).AddYears(RelativeValue);
                    if (FirstOrLast == "Last")
                    {
                        ret = ret.AddYears(1).AddDays(-1);
                    }
                    break;
            }
            return ret;
        }

        public static int DaysInYear(int year)
        {
            DateTime firstdate = new DateTime(year, 1, 1);
            DateTime lastdate = new DateTime(year, 12, 31);
            return (lastdate - firstdate).Days + 1;
        }

        public static int DaysInQuarter(DateTime date)
        {
            int quarterNumber = (date.Month - 1) / 3 + 1;
            DateTime firstDayOfQuarter = new DateTime(date.Year, (quarterNumber - 1) * 3 + 1, 1);
            DateTime lastDayOfQuarter = firstDayOfQuarter.AddMonths(3).AddDays(-1);
            return (lastDayOfQuarter - firstDayOfQuarter).Days + 1;
        }

        public static int DayInQuarter(DateTime date)
        {
            int quarterNumber = (date.Month - 1) / 3 + 1;
            DateTime firstDayOfQuarter = new DateTime(date.Year, (quarterNumber - 1) * 3 + 1, 1);
            return (date - firstDayOfQuarter).Days + 1;
        }

        public static bool CompareObjects(object obj1, object obj2)
        {
            if (obj1 != null && obj2 != null)
            {
                if (obj1.GetType().Equals(obj2.GetType()))
                {
                    return obj1.Equals(obj2);
                }
                else return false;
            }

            if (obj1 == null && obj2 == null) return true;
            return false;

        }

        public static string ReverseString(string s)
        {
            char[] arr = s.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }

        public static string PrintArray(double?[] arr)
        {
            string ret = "";
            for (int i = 0; i < arr.Length; i++) { ret += i.ToString() + " " + arr[i].ToString(); }
            return ret;
        }

        public static Dictionary<object, object> StringToDictionary(string content, IdList keyColumnIndexes, IdList valueColumnIndexes)
        {
            // this method converts Tab Delimited text into a dictionary
            Dictionary<object, object> ret = new Dictionary<object, object>();
            if (content.Trim() == "") return ret;

            string[] lines = content.Split("\n".ToCharArray());

            // first line is HEAD line, skip
            if (lines.Length == 1) return ret;

            // separate by Tab
            string[] line;

            // process from 2nd line
            for (int i = 1; i < lines.Length; i++)
            {
                // King Sun 2014-08-24 use TrimEnd otherwise trailing Tab can be gone
                //line = lines[i].Trim().Split("\t".ToCharArray());
                line = lines[i].TrimEnd(GeneralSettings.trimChars).Split("\t".ToCharArray());
                string key = "";
                string value = "";
                foreach (int colIndex in keyColumnIndexes)
                {
                    if (key != "") key += IdDelimiterApp;
                    key += line[colIndex];
                }
                foreach (int colIndex in valueColumnIndexes)
                {
                    if (value != "") value += IdDelimiterApp;
                    value += line[colIndex];
                }
                ret.Add(key, value);
            }

            return ret;
        }

        public static Dictionary<string, int> StringToDictionary(string content, IdList keyColumnIndexes, int valueColumnIndex)
        {
            // this method converts Tab Delimited text into a dictionary
            Dictionary<string, int> ret = new Dictionary<string, int>();
            if (content.Trim() == "") return ret;

            string[] lines = content.Split("\n".ToCharArray());

            // first line is HEAD line, skip
            if (lines.Length == 1) return ret;

            // separate by Tab
            string[] line;

            // process from 2nd line
            for (int i = 1; i < lines.Length; i++)
            {
                // King Sun 2014-08-24 use TrimEnd otherwise trailing Tab can be gone
                //line = lines[i].Trim().Split("\t".ToCharArray());
                line = lines[i].TrimEnd(GeneralSettings.trimChars).Split("\t".ToCharArray());
                string key = "";
                int value = 0;
                foreach (int colIndex in keyColumnIndexes)
                {
                    if (key != "") key += IdDelimiterApp;
                    key += line[colIndex];
                }
                value = Convert.ToInt32(line[valueColumnIndex]);
                ret.Add(key, value);
            }

            return ret;
        }
    }

    public class SysDataType
    {
        public enum DataType
        {
            TypeString = 1,
            TypeNumber = 2,
            TypeBoolean = 3,
            TypeDate = 4//,
            //TypeUnknown = 5
            //TypeBOM = 10
            //TypeElement = 20
            //TypeFcstUnit = 30
        }

        public static System.Type ToSystemType(SysDataType dt)
        {
            switch (dt.Id)
            {
                case (DataType.TypeString): return typeof(string);
                case (DataType.TypeDate): return typeof(DateTime);
                case (DataType.TypeNumber): return typeof(double);
                case (DataType.TypeBoolean): return typeof(bool);
                default: return typeof(string);
            }
        }

        private DataType m_Id = DataType.TypeString;
        public SysDataType()
        {
        }

        public SysDataType(DataType Id)
        {
            this.Id = Id;
        }

        public SysDataType(string Name)
        {
            DataType[] types = Enum.GetValues(typeof(DataType)) as DataType[];
            foreach (DataType i in types)
            {
                if (new SysDataType(i).Name == Name)
                {
                    this.Id = i;
                    break;
                }
            }
        }

        public DataType Id
        {
            get { return m_Id; }
            set { m_Id = value; }
        }

        public string Name
        {
            get
            {
                switch (Id)
                {
                    case DataType.TypeString:
                        return "STRING";
                    case DataType.TypeNumber:
                        return "NUMBER";
                    case DataType.TypeBoolean:
                        return "BOOLEAN";
                    case DataType.TypeDate:
                        return "DATE";
                    //Case DataType.TypeBOM
                    //    Return "BOM"
                    default:
                        return "STRING";
                }
            }
        }

        public static double StringToDouble(string Value)
        {
            double ret = 0;
            if (!string.IsNullOrEmpty(Value) && GeneralSettings.IsNumeric(Value))
                ret = Convert.ToDouble(Value, GeneralSettings.DefCulture);

            return ret;
        }

        public static int StringToInt(string Value)
        {
            int ret = 0;
            if (!string.IsNullOrEmpty(Value) && GeneralSettings.IsNumeric(Value))
                ret = Convert.ToInt32(Value, GeneralSettings.DefCulture);

            return ret;
        }

        public static DateTime StringToDate(string Value)
        {
            DateTime ret = GeneralSettings.MinDate;
            if (!string.IsNullOrEmpty(Value))
            {
                ret = DateTime.ParseExact(Value, GeneralSettings.DefDateTimeFormat, System.Globalization.CultureInfo.InvariantCulture);
            }
            return ret;
        }

        public static DateTime StringToDateTime(string value)
        {
            DateTime ret = GeneralSettings.MinDateTime;
            if (!string.IsNullOrEmpty(value))
            {
                ret = DateTime.ParseExact(value, GeneralSettings.DefDateTimeFormat, CultureInfo.InvariantCulture);
            }
            return ret;
        }

        public static object StringToObject(SysDataType.DataType DataType, string Value)
        {
            object ret = null;
            switch (DataType)
            {
                case SysDataType.DataType.TypeString:
                    ret = Value;
                    break;
                case SysDataType.DataType.TypeDate:
                    if (!string.IsNullOrEmpty(Value))
                    {
                        ret = DateTime.ParseExact(Value, GeneralSettings.DefDateTimeFormat, System.Globalization.CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        // King Sun 2013-10-14 if String is empty, then return nothing as Date
                        //ret = GeneralSettings.minDate;
                    }
                    break;
                case SysDataType.DataType.TypeBoolean:
                    if (!string.IsNullOrEmpty(Value))
                        ret = Convert.ToBoolean(Value);
                    break;
                case SysDataType.DataType.TypeNumber:
                    if (!string.IsNullOrEmpty(Value))
                    {
                        try
                        {
                            ret = Math.Round(Convert.ToDouble(Value), 6);
                        }
                        catch (Exception)
                        {
                            ret = Value;
                        }
                    }
                    break;
            }
            return ret;
        }

        public static string DateToString(DateTime Value)
        {
            return ObjectToString(SysDataType.DataType.TypeDate, Value, GeneralSettings.DefCulture);
        }


        public static string ObjectToString(object Value)
        {
            // This function is to convert local formatted values to standard (Englished Based)
            if (Value.GetType().Equals(typeof(decimal)))
            {
                return Convert.ToString(Value);
            }
            else if (Value.GetType().Equals(typeof(DateTime)))
            {
                DateTime dt = Convert.ToDateTime(Value);
                return dt.ToString(GeneralSettings.DefDateTimeFormat);
            }
            else
            {
                return Convert.ToString(Value);
            }
        }

        public static string ObjectToString(SysDataType.DataType DataType, object Value)
        {
            return ObjectToString(DataType: DataType, Value: Value, cul: GeneralSettings.DefCulture);
        }

        public static string ObjectToString(SysDataType.DataType DataType, object Value, CultureInfo cul)
        {
            if (Value == null)
                return "";

            if (Value.Equals(DBNull.Value))
                return "";

            if (Value.GetType() == typeof(string) && Value.ToString() == "")
                return "";

            string ret = "";
            switch (DataType)
            {
                case SysDataType.DataType.TypeString:
                    ret = Value.ToString();
                    break;
                case SysDataType.DataType.TypeDate:
                    ret = Convert.ToDateTime(Value).ToString(GeneralSettings.DefDateTimeFormat);
                    break;
                case SysDataType.DataType.TypeBoolean:
                    ret = Math.Abs(Convert.ToInt32(Value)).ToString();
                    break;
                case SysDataType.DataType.TypeNumber:
                    if (Value.GetType() == typeof(string) && Value.ToString().Contains(cul.NumberFormat.NumberGroupSeparator))
                        // The C# Convert.ToDecimal method handles group separators different
                        // than Excel. We need to handle them here to avoid counter-intuitive results.
                        //    1,770 -> 1770  OK
                        //    1,77  -> 177   NOT OK, Excel interpretes it as an invalid number.
                        Logging.Log("Group separator not accepted in numbers: " + Value.ToString() + "\n");
                    else
                    {
                        try
                        {
                            ret = Math.Round(Convert.ToDecimal(Value, cul), 6).ToString(cul); ;
                        }
                        catch (Exception ex)
                        {
                            Logging.Log("Fail to convert to decimal for value :" + Value.ToString() + "\n" + ex.Message);
                        }
                    }
                    break;
                //case AttributeDataType.DataType.TypeUnknown: {
                //        // driven by object type
                //        return ObjectToString(ToAttributeDataType(Value), Value);
                //    }
            }
            return ret;
        }
    }

    public class DbConversion
    {
        public class QueryParameter
        {
            public string Name;

            public object Value;
            // empty constructor

            public QueryParameter()
            {
            }
            // constructor with direct input
            public QueryParameter(string Name, object Value)
            {
                this.Name = Name;
                this.Value = Value;
            }
        }

        public class QueryTarget
        {
            public string QueryString = "";

            public List<QueryParameter> Parameters = new List<QueryParameter>();
            // empty constructor

            public QueryTarget()
            {
            }

            // constructor with direct input
            public QueryTarget(string sqlStr)
            {
                QueryString = sqlStr;
            }

            // constructor with direct input
            public QueryTarget(string sqlStr, List<QueryParameter> ParameterList)
            {
                QueryString = sqlStr;
                Parameters = ParameterList;
            }

            public string PrintString()
            {
                string ret = QueryString;
                foreach (QueryParameter para in Parameters)
                {
                    ret += "\n" + para.Name + ": " + para.Value.ToString();
                }
                ret += "\n";
                return ret;
            }

        }

        public static string RepeatQuestionMark(int repeatTimes)
        {
            string ret = "";
            for (int i = 0; i < repeatTimes; i++)
            {
                if (i == 0) ret += "?";
                else ret += ",?";
            }

            return ret;
        }

        public static string idListToSQL(string idList)
        {
            // this function gets an id list delimited by : and convert to SQL statement as 
            // SELECT id_1 FROM DUAL UNION ALL SELECT id_2 FROM DUAL

            string[] ids = null;
            int i = 0;
            string sqlStr = "";
            ids = idList.Split(GeneralSettings.IdDelimiterApp.ToCharArray(0, 1));
            for (i = 0; i <= ids.Length - 1; i++)
            {
                if (GeneralSettings.IsNumeric(ids[i]))
                {
                    if (!string.IsNullOrEmpty(sqlStr))
                        sqlStr = sqlStr + " UNION ALL ";
                    sqlStr = sqlStr + "SELECT " + ids[i] + " FROM DUAL";
                }
            }

            if (string.IsNullOrEmpty(sqlStr))
                sqlStr = "SELECT 0 FROM DUAL";

            return sqlStr;
        }

        public static string idListToInClause(string ColumnName, List<int> idList, bool NotIn = false)
        {
            // this function gets an IN list delimited by : and convert to SQL statement as 
            // when id count > 1000 then use OR to link the following ids
            // (COLUMN_NAME IN (id_1, id_2, ...) OR COLUMN_NAME IN (id_1001, id_1002, ...))

            int cnt = 0;
            string inStr = "";
            List<string> inList = new List<string>();

            if (idList.Count > 1)
            {
                foreach (int id in idList)
                {
                    if (cnt > 0)
                        inStr += GeneralSettings.IdDelimiterDb;
                    inStr += id.ToString();
                    cnt = cnt + 1;
                    if (cnt == 1000)
                    {
                        inList.Add(inStr);
                        inStr = "";
                        cnt = 0;
                    }
                }
                if (!string.IsNullOrEmpty(inStr))
                {
                    inList.Add(inStr);
                }
            }
            else if (idList.Count == 1)
            {
                if (NotIn)
                {
                    return "(" + ColumnName + "<>" + idList[0].ToString() + ")";
                }
                else
                {
                    return "(" + ColumnName + "=" + idList[0].ToString() + ")";
                }

            }
            else
            {
                inList.Add("0");
            }

            string sqlStr = "";
            for (int i = 0; i <= inList.Count - 1; i++)
            {
                if (NotIn)
                {
                    if (i > 0)
                        sqlStr += " AND ";
                    sqlStr += ColumnName + " NOT IN (" + inList[i] + ")";
                }
                else
                {
                    if (i > 0)
                        sqlStr += " OR ";
                    sqlStr += ColumnName + " IN (" + inList[i] + ")";
                }
            }

            sqlStr = "(" + sqlStr + ")";

            return sqlStr;
        }

        public static string idListToInClause(string ColumnName, string idList, bool NotIn = false)
        {
            // this function gets an IN list delimited by : and convert to SQL statement as 
            // when id count > 1000 then use OR to link the following ids
            // (COLUMN_NAME IN (id_1, id_2, ...) OR COLUMN_NAME IN (id_1001, id_1002, ...))

            List<int> ids = new List<int>();
            List<string> inList = new List<string>();

            foreach (String id in idList.Split(GeneralSettings.IdDelimiterApp.ToCharArray(0, 1)))
            {
                if (GeneralSettings.IsNumeric(id))
                {
                    ids.Add(Convert.ToInt32(id));
                }
            }

            return idListToInClause(ColumnName: ColumnName, idList: ids, NotIn: NotIn);
        }
    }

    public class EnumObject
    {
        public Type EnumType { get; set; }
        public int Id { get; set; }
        //public string Description { get; set; }
        //public string Category { get; set; }

        public EnumObject(Type EnumType, int Id)
        {
            this.EnumType = EnumType;
            this.Id = Id;
            //this.Description = EnumHelper.GetDescription((typeof(EnumType))Id);
        }

        public string Description
        {
            get
            {
                string ret = Name;
                MemberInfo[] memInfo = EnumType.GetMember(ret);

                if (memInfo != null && memInfo.Length > 0)
                {
                    object[] attrs = memInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);

                    if (attrs != null && attrs.Length > 0)
                    {
                        return ((DescriptionAttribute)attrs[0]).Description;
                    }
                }

                return ret;
            }
        }

        public string FullDescription
        {
            get
            {
                string ret = this.Category;
                if (ret != "") ret += ": ";
                ret += this.Description;
                return ret;
            }
        }

        public string Category
        {
            get
            {
                string ret = "";
                MemberInfo[] memInfo = EnumType.GetMember(Name);

                if (memInfo != null && memInfo.Length > 0)
                {
                    object[] attrs = memInfo[0].GetCustomAttributes(typeof(CategoryAttribute), false);

                    if (attrs != null && attrs.Length > 0)
                    {
                        return ((CategoryAttribute)attrs[0]).Category;
                    }
                }

                return ret;
            }
        }

        public string Name
        {
            get
            {
                try
                {
                    return Enum.GetName(EnumType, Id);
                }
                catch
                {
                    return "Invalid Id: " + Id.ToString();
                }
            }
        }
    }

    public class EnumObjectList : List<EnumObject>
    {
        public EnumObjectList()
        {
        }

        public EnumObjectList(Type EnumType)
        {
            int[] ids = Enum.GetValues(enumType: EnumType) as int[];
            foreach (int id in ids)
            {
                this.Add(new EnumObject(EnumType: EnumType, Id: id));
            }
        }

        public EnumObjectList(Type EnumType, string Category)
        {
            int[] ids = Enum.GetValues(enumType: EnumType) as int[];
            foreach (int id in ids)
            {
                EnumObject obj = new EnumObject(EnumType: EnumType, Id: id);
                if (obj.Category == Category) this.Add(obj);
            }
        }

        public EnumObject GetElement(int id)
        {
            return this.Find(e => e.Id == id);
        }

        public EnumObject GetElement(string Name)
        {
            return this.Find(e => e.Name == Name);
        }
    }

    public static class EnumHelper
    {
        /// <summary>
        /// Retrieve the description on the enum, e.g.
        /// [Description("Bright Pink")]
        /// BrightPink = 2,
        /// Then when you pass in the enum, it will retrieve the description
        /// </summary>
        /// <param name="en">The Enumeration</param>
        /// <returns>A string representing the friendly name</returns>
        public static string GetDescription(Enum en)
        {
            Type type = en.GetType();

            MemberInfo[] memInfo = type.GetMember(en.ToString());

            if (memInfo != null && memInfo.Length > 0)
            {
                object[] attrs = memInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);

                if (attrs != null && attrs.Length > 0)
                {
                    return ((DescriptionAttribute)attrs[0]).Description;
                }
            }

            return en.ToString();
        }

        public static string GetCategory(Enum en)
        {
            Type type = en.GetType();

            MemberInfo[] memInfo = type.GetMember(en.ToString());

            if (memInfo != null && memInfo.Length > 0)
            {
                object[] attrs = memInfo[0].GetCustomAttributes(typeof(CategoryAttribute), false);

                if (attrs != null && attrs.Length > 0)
                {
                    return ((CategoryAttribute)attrs[0]).Category;
                }
            }

            return en.ToString();
        }

        public static int GetEnumByDescription(Type enumType, string Description)
        {
            foreach (FieldInfo field in enumType.GetFields())
            {
                var attribute = System.Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) as DescriptionAttribute;
                if (attribute != null)
                {
                    if (attribute.Description == Description)
                    {
                        return (int)field.GetValue(null);
                    }
                }
                else
                {
                    if (field.Name == Description)
                    {
                        return (int)field.GetValue(null);
                    }
                }
            }
            return 0;
        }
    }
}
