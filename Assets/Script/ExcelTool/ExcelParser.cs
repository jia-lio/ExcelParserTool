using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using UnityEditor;
using UnityEngine;
using Assembly = System.Reflection.Assembly;

namespace Tool.Excel
{
    public static class ExcelParser
    {
        public static void ParseExcel(string excelPath, string tablePath, string assetPath)
        {
            var fullPath = Path.Combine(Application.dataPath.Substring(0, Application.dataPath.Length - 6), excelPath);
            
            using var workbook = new XLWorkbook(fullPath);
            var worksheet = workbook.Worksheet(1);

            var headerRow = worksheet.Row(1);
            var typeRow = worksheet.Row(2);

            var columnCount = worksheet.LastColumnUsed().ColumnNumber();
            var rowCount = worksheet.LastRowUsed().RowNumber();

            bool[] hasDataInColumn = new bool[columnCount + 1];

            for (var col = 1; col <= columnCount; col++)
            {
                for (var rowIdx = 3; rowIdx <= rowCount; rowIdx++)
                {
                    var val = worksheet.Row(rowIdx).Cell(col).GetString();
                    if (string.IsNullOrEmpty(val) == false)
                    {
                        hasDataInColumn[col] = true;
                        break;
                    }
                }
            }

            CreateTableCS(excelPath, headerRow, typeRow, tablePath);
            CreateScriptableObjectCS(excelPath, tablePath);
        }

        private static void CreateTableCS(string excelPath, IXLRow headerRow, IXLRow typeRow, string tablePath)
        {
            var csName = Path.GetFileNameWithoutExtension(excelPath) + "Table";
            var usings = $"using System;\nusing UnityEngine;\n\n";
            var serializable = "[Serializable]\n";
            var code = usings + serializable + $"public class {csName} : BaseTable\n{{\n";
            var columnCount = headerRow.Worksheet.LastColumnUsed().ColumnNumber();

            for (var i = 1; i <= columnCount; i++)
            {
                var fieldName = headerRow.Cell(i).GetString();
                var fieldTypeString = typeRow.Cell(i).GetString().ToLowerInvariant();
                var type = ConvertToCSharpType(fieldTypeString);

                if (string.IsNullOrEmpty(fieldName) || string.IsNullOrEmpty(type))
                    continue;

                if (fieldName.Equals("ID", StringComparison.OrdinalIgnoreCase))
                    continue;

                code += $"    public {type} {fieldName};\n";
            }

            code += "}\n";

            var fullDirectoryPath = Path.Combine("Assets", tablePath);
            var scriptPath = Path.Combine(tablePath, $"{csName}.cs");
            Directory.CreateDirectory(fullDirectoryPath);
            File.WriteAllText($"Assets/{scriptPath}", code);
            
#if UNITY_EDITOR
            AssetDatabase.Refresh();
#endif
        }

        private static void CreateScriptableObjectCS(string excelPath, string tablePath)
        {
            var tableName = Path.GetFileNameWithoutExtension(excelPath);
            var tableCSName = $"{tableName}Table";
            var csName = $"{tableName}TableSO";
            var usings = "using UnityEngine;\n\n";
            var serializable = "[Serializable]\n";
            var attr = $"[CreateAssetMenu(fileName = \"{csName}\", menuName = \"Table/{tableName} Table\")]\n";
            var code = usings + attr + $"public class {csName} : BaseScriptableObject<{tableCSName}>{{ }}\n";
            
            var fullDirectoryPath = Path.Combine("Assets", tablePath);
            var scriptPath = Path.Combine(tablePath, $"{csName}.cs");
            Directory.CreateDirectory(fullDirectoryPath);
            File.WriteAllText($"Assets/{scriptPath}", code);
            
#if UNITY_EDITOR
            AssetDatabase.Refresh();
#endif
        }

        public static void CreateScriptableObjectAsset(string excel, string assetPath)
        {
            var tableName = Path.GetFileNameWithoutExtension(excel);
            var tableCSName = $"{tableName}Table";
            var soCSName = $"{tableCSName}SO";

            var assembly = Assembly.Load("Assembly-CSharp");
            var tableType = assembly.GetTypes().FirstOrDefault(t => t.Name == tableCSName);
            var soType = assembly.GetTypes().FirstOrDefault(t => t.Name == soCSName);

            if(tableType == null || soType == null)
                return;

            var so = ScriptableObject.CreateInstance(soType);
            var fullPath = Path.Combine(Application.dataPath.Substring(0, Application.dataPath.Length - 6), excel);
            using var workbook = new XLWorkbook(fullPath);
            var worksheet = workbook.Worksheet(1);

            var headerRow = worksheet.Row(1);
            var typeRow = worksheet.Row(2);
            var rowCount = worksheet.LastRowUsed().RowNumber();
            var columnCount = worksheet.LastColumnUsed().ColumnNumber();

            var rowsField = soType.GetField("rows");
            if (rowsField == null)
                return;

            var rowsList = rowsField.GetValue(so);
            if (rowsList == null)
            {
                var listType = typeof(List<>).MakeGenericType(tableType);
                rowsList = Activator.CreateInstance(listType);
                rowsField.SetValue(so, rowsList);
            }

            var addMethod = rowsList.GetType().GetMethod("Add");
            var tableProperties = tableType.GetProperties(BindingFlags.Public | BindingFlags.Instance);

            for (var rowIndex = 3; rowIndex <= rowCount; rowIndex++)
            {
                var rowObj = Activator.CreateInstance(tableType);

                for (var col = 1; col <= columnCount; col++)
                {
                    var propName = headerRow.Cell(col).GetString();
                    if (string.IsNullOrEmpty(propName))
                        continue;

                    var tableFields = tableType.GetFields(BindingFlags.Public | BindingFlags.Instance);
                    var prop = tableProperties.FirstOrDefault(p => p.Name.Equals(propName, StringComparison.OrdinalIgnoreCase));
                    var field = tableFields.FirstOrDefault(f => f.Name.Equals(propName, StringComparison.OrdinalIgnoreCase));
                    
                    if (prop == null && field == null)
                        continue;

                    var cellValue = worksheet.Row(rowIndex).Cell(col).GetString();
                    if (string.IsNullOrEmpty(cellValue))
                        continue;

                    try
                    {
                        object convertedValue;
                        if (prop != null)
                        {
                            convertedValue = ConvertCellValue(cellValue, prop.PropertyType);
                            prop.SetValue(rowObj, convertedValue);
                        }
                        else
                        {
                            convertedValue = ConvertCellValue(cellValue, field.FieldType);
                            field.SetValue(rowObj, convertedValue);
                        }
                    }
                    catch (Exception e)
                    {
                        Debug.LogWarning($"Cell 변환 실패: {cellValue} to {prop.PropertyType.Name} ({e.Message})");
                    }
                }

                addMethod.Invoke(rowsList, new[] { rowObj });
            }

            var assetName = soCSName + ".asset";
            var path = Path.Combine(assetPath, assetName).Replace("\\", "/");
            Directory.CreateDirectory(assetPath);

            var existingAsset = AssetDatabase.LoadAssetAtPath(path, soType);
            if (existingAsset != null)
            {
                AssetDatabase.DeleteAsset(path);
            }
            
            AssetDatabase.CreateAsset(so, path);
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();

            Debug.Log($"ScriptableObject 에셋 생성 완료: {path}");
        }

        private static string ConvertToCSharpType(string typeStr)
        {
            return typeStr switch
            {
                "int" => "int",
                "float" => "float",
                "bool" => "bool",
                "string" => "string",
                "datetime" => "DateTime",
                "vector3" => "Vector3",
                _ => null
            };
        }
        
        private static object ConvertCellValue(string cellValue, Type fieldType)
        {
            try
            {
                if (fieldType == typeof(int))
                    return int.Parse(cellValue);
                if (fieldType == typeof(float))
                    return float.Parse(cellValue, System.Globalization.CultureInfo.InvariantCulture);
                if (fieldType == typeof(bool))
                    return bool.Parse(cellValue);
                if (fieldType == typeof(string))
                    return cellValue;
                if (fieldType == typeof(DateTime))
                    return DateTime.Parse(cellValue);
                if (fieldType == typeof(Vector3))
                {
                    var trimmed = cellValue.Trim('(', ')').Replace(" ", "");
                    var parts = trimmed.Split(',');
                    return new Vector3(
                        float.Parse(parts[0]),
                        float.Parse(parts[1]),
                        float.Parse(parts[2])
                    );
                }
            }
            catch (Exception e)
            {
                Debug.LogWarning($"Cell 변환 실패: {cellValue} to {fieldType.Name} ({e.Message})");
            }

            return null;
        }
    }
}
