using System.Collections.Generic;
using mohism.excel.exporter;
using UnityEditor;
using UnityEngine;

namespace mohism.excel.editor {
    /// <summary>
    /// excel轉檔工具
    /// </summary>
    public class ExcelTool : EditorWindow {
        /// <summary>
        /// 單例
        /// </summary>
        private static ExcelTool _inst = null;

        /// <summary>
        /// excel文件列表
        /// </summary>
        private static List<string> _files = new List<string>();

        /// <summary>
        /// 檔案總數
        /// </summary>
        private static int _count { get { return _files.Count; } }

        /// <summary>
        /// 輸出路徑
        /// </summary>
        private static string _dest = string.Empty;

        /// <summary>
        /// 紀錄本地輸出目錄
        /// </summary>
        private const string DEST_KEY = "dest";

        /// <summary>
        /// 顯示視窗
        /// </summary>
        [MenuItem("Assets/表格轉檔")]
        private static void ShowForm() {
            _inst = GetWindow<ExcelTool>();

            if (PlayerPrefs.HasKey(DEST_KEY)) {
                _dest = PlayerPrefs.GetString(DEST_KEY);
            }
            else {
                _dest = Application.dataPath;
            }

            load();

            _inst.Show();
        }

        /// <summary>
        /// 讀取文件
        /// </summary>
        private static void load() {
            _files.Clear();

            var selects = (object[])Selection.objects;

            // 沒有選到的文件
            if (selects.Length <= 0) {
                return;
            }

            foreach (var elm in selects) {
                var path = AssetDatabase.GetAssetPath((Object)elm);

                if (path.EndsWith(".xlsx")) {
                    _files.Add(path);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void OnGUI() {
            // 顯示輸出路徑
            EditorGUILayout.LabelField(string.Format("輸出路徑: {0}", _dest));

            // 修改輸出路徑
            if (GUILayout.Button("修改輸出路徑")) {
                _dest = EditorUtility.OpenFolderPanel("選擇輸出路徑", _dest, "");
                PlayerPrefs.SetString(DEST_KEY, _dest);
            }

            // 繪製文件列表
            drawFiles();

            // 執行
            if (GUILayout.Button("執行")) {
                execute();
            }
        }

        /// <summary>
        /// 繪製文件列表
        /// </summary>
        private void drawFiles() {
            if (_files.Count <= 0) {
                EditorGUILayout.LabelField("無已選的excel");
                return;
            }

            EditorGUILayout.LabelField(string.Format("共{0}個excel已選, 如下列:", _count));

            GUILayout.BeginVertical();
            GUILayout.BeginScrollView(Vector2.zero, false, true, GUILayout.Height(250));

            // 繪製文件列表
            for (var i = 0; i < _count; i++) {
                GUILayout.BeginHorizontal();
                GUILayout.Label(string.Format("{0}.{1}", i, _files[i]));
                GUILayout.EndHorizontal();
            }

            GUILayout.EndScrollView();
            GUILayout.EndVertical();
        }

        /// <summary>
        /// 執行轉檔
        /// </summary>
        private void execute() {
            // 未設定輸出路徑
            if (string.IsNullOrEmpty(_dest)) {
                Debug.LogError("轉檔失敗, 未設定輸出路徑");
                return;
            }

            // 清除log
            Debug.ClearDeveloperConsole();

            Debug.LogFormat("開始轉檔, 共{0}個excel參與", _count);

            for (int i = 0; i < _count; i++) {
                Debug.LogFormat("{0}.{1}", i, _files[i]);

                // 輸出文件
                var table = new ParseTable(_files[i]);
                table.export(_dest, new JsonExporterC(), new JsonExporterS());

                // 刷新資源顯示
                AssetDatabase.Refresh();
            }

            Debug.Log("轉檔完成");

            _inst.Close();
        }
    }
}
