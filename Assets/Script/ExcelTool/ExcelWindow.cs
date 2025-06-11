using System.Collections.Generic;
using System.IO;
using UnityEditor;
using UnityEngine;

namespace Tool.Excel
{
    public class ExcelWindow : EditorWindow
    {
        private string folderPath = "Assets/Data/GameDesign/Excel";
        private string scriptableObjectPath = "Assets/Data/GameDesign/ScriptableObject";
        private string tablePath = "Data/GameDesign/Table";
        private List<ExcelEntry> excelEndtries = new();

        [MenuItem("Tools/Excel", priority = 0)]
        private static void Show()
        {
            var window = (ExcelWindow)GetWindow(typeof(ExcelWindow), false, "Excel Tool");
            window.minSize = new Vector2(650, 500);
            window.RefreshExcelList();
        }

        private void RefreshExcelList()
        {
            excelEndtries.Clear();

            string[] allExcel = AssetDatabase.FindAssets("t:DefaultAsset", new[] { folderPath });
            foreach (var excel in allExcel)
            {
                var path = AssetDatabase.GUIDToAssetPath(excel);
                if (path.EndsWith(".xlsx") == false)
                    continue;

                var entry = new ExcelEntry
                {
                    excelPath = path
                };

                var fileNameWithoutExt = Path.GetFileNameWithoutExtension(path);
                var soPath = $"{scriptableObjectPath}/{fileNameWithoutExt}TableSO.asset";

                var so = AssetDatabase.LoadAssetAtPath<ScriptableObject>(soPath);
                if (so != null)
                {
                    entry.assetPath = soPath;
                }
                
                excelEndtries.Add(entry);
            }
        }
        
        private void CreateTable()
        {
            foreach (var excel in excelEndtries)
            {
                ExcelParsing(excel);
            }
        }

        private void ExcelParsing(ExcelEntry excel)
        {
            ExcelParser.ParseExcel(excel.excelPath, tablePath, scriptableObjectPath);
        }
        
        private void OnGUI()
        {
            //경로
            GUILayout.Space(10);
            EditorGUILayout.LabelField("엑셀 파일 경로", EditorStyles.boldLabel);

            EditorGUILayout.BeginHorizontal();

            EditorGUI.BeginDisabledGroup(true);
            EditorGUILayout.TextField(folderPath);
            EditorGUI.EndDisabledGroup();

            if (GUILayout.Button("...", GUILayout.Width(30)))
            {
                var newPath = EditorUtility.OpenFolderPanel("엑셀 폴더 선택", Application.dataPath, "");
                if (string.IsNullOrEmpty(newPath) == false && newPath.StartsWith(Application.dataPath))
                {
                    folderPath = "Assets" + newPath.Substring(Application.dataPath.Length);
                }
            }

            EditorGUILayout.EndHorizontal();

            //새로고침
            if (GUILayout.Button("엑셀 목록 새로고침"))
            {
                RefreshExcelList();
            }
            
            //CS 생성
            if (GUILayout.Button("CS 생성"))
            {
                CreateTable();
                RefreshExcelList();
            }
            
            //SO 생성
            if (GUILayout.Button("SO 생성"))
            {
                foreach (var excel in excelEndtries)
                {
                    ExcelParser.CreateScriptableObjectAsset(excel.excelPath, scriptableObjectPath);
                }
                
                RefreshExcelList();
            }

            //파일에 있는 엑셀 보여주기
            if (excelEndtries.Count == 0)
            {
                EditorGUILayout.LabelField("파일이 없습니다.");
            }
            else
            {
                foreach (var excel in excelEndtries)
                {
                    EditorGUILayout.BeginHorizontal();

                    var excelName = Path.GetFileName(excel.excelPath);
                    if (GUILayout.Button(excelName, GUILayout.Width(300)))
                    {
                        EditorUtility.RevealInFinder(excel.excelPath);
                    }

                    if (string.IsNullOrEmpty(excel.assetPath) == false)
                    {
                        var assetName = Path.GetFileNameWithoutExtension(excel.assetPath);

                        if (GUILayout.Button(assetName, GUILayout.Width(300)))
                        {
                            Selection.activeObject = AssetDatabase.LoadAssetAtPath<ScriptableObject>(excel.assetPath);
                        }
                    }
                    else
                    {
                        EditorGUILayout.LabelField("없음");
                    }
                    
                    EditorGUILayout.EndHorizontal();
                }
            }
        }
    }
}