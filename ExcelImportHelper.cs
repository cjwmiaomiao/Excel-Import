
namespace SPATool
{
    using OfficeOpenXml;
    using Serenity;
    using Serenity.Data;
    using Serenity.Services;
    using Serenity.Web;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using LicenseContext = OfficeOpenXml.LicenseContext;

    public static class ExcelImportHelper
    {
        public static List<T> ExcelImport<T>(IDbConnection connection, ExcelImportRequest request)
        {
            var response = new List<T>();

            request.CheckNotNull();
            Check.NotNullOrWhiteSpace(request.FileName, "filename");
            UploadHelper.CheckFileNameSecurity(request.FileName);

            if (!request.FileName.StartsWith("temporary/"))
                throw new ArgumentOutOfRangeException("filename");

            // Load Excel file using EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage ep = new ExcelPackage();
            using (var fs = new FileStream(UploadHelper.DbFilePath(request.FileName), FileMode.Open, FileAccess.Read))
                ep.Load(fs);

            var sheet = ep.Workbook.Worksheets[0];
            int rowNum = sheet.Dimension.Rows;     // Total number of rows
            int cellNum = sheet.Dimension.Columns; // Total number of columns

            // Initialize list to store objects of type T
            List<T> excelList = new List<T>();

            PropertyInfo[] properties = typeof(T).GetProperties();
            Dictionary<string, PropertyInfo> propertyMap = new Dictionary<string, PropertyInfo>();
            foreach (PropertyInfo property in properties)
            {
                var displayNameAttr = (DisplayNameAttribute)property.GetCustomAttribute(typeof(DisplayNameAttribute));
                if (displayNameAttr != null)
                {
                    propertyMap.Add(displayNameAttr.DisplayName, property);
                }
            }

            for (int i = 2; i <= rowNum; i++)
            {
                bool isEmptyRow = true;
           
                T newPojo = Activator.CreateInstance<T>();

                foreach (var kvp in propertyMap)
                {
                    string displayName = kvp.Key;
                    PropertyInfo property = kvp.Value;

                    int columnIndex = -1;
                    for (int j = 1; j <= cellNum; j++) 
                    {
                        string cellValue = sheet.Cells[1, j].Value?.ToString(); // Header row 
                        if (cellValue == displayName)
                        {
                            columnIndex = j;
                            break;
                        }
                    }

                    if (columnIndex != -1)
                    {
                        Type type = property.PropertyType;
                        object cellValue = sheet.Cells[i, columnIndex].Value;

                        if (cellValue!=null&&!string.IsNullOrWhiteSpace(cellValue.ToString()))
                        {
                            isEmptyRow = false;
                        }

                        if (type == typeof(int))
                        {
                            int value;
                            if (cellValue != null && int.TryParse(cellValue.ToString(), out value))
                            {
                                property.SetValue(newPojo, value);
                            }
                        }
                        else if (type == typeof(long))
                        {
                            long value;
                            if (cellValue != null && long.TryParse(cellValue.ToString(), out value))
                            {
                                property.SetValue(newPojo, value);
                            }
                        }
                        else if (type == typeof(string))
                        {
                            if (cellValue != null)
                            {
                                property.SetValue(newPojo, cellValue.ToString());
                            }
                        }
                        else if (type == typeof(DateTime))
                        {
                            DateTime value;
                            if (cellValue != null && DateTime.TryParse(cellValue.ToString(), out value))
                            {
                                property.SetValue(newPojo, value);
                            }
                        }
                        else if (type == typeof(decimal))
                        {
                            decimal value;
                            if (cellValue != null && decimal.TryParse(cellValue.ToString(), out value))
                            {
                                property.SetValue(newPojo, value);
                            }
                        }
                    }
                }

                if (isEmptyRow==true)
                {
                    break;
                }
                excelList.Add(newPojo);
            }

            return excelList;
        }

    }
}