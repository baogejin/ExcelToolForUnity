using Newtonsoft.Json;
using System.Collections.Generic;
using UnityEngine;
using YooAsset;

namespace GameData
{
    public class ExampleInfo
    {
        /// <summary>
        /// 数字
        /// </summary>
        public int ID;
        /// <summary>
        /// 字符串
        /// </summary>
        public string Name;
        /// <summary>
        /// 服务器不用字段
        /// </summary>
        public float Ignore;
        /// <summary>
        /// 数组
        /// </summary>
        public List<int> Slc1;
        /// <summary>
        /// 数组
        /// </summary>
        public List<float> Slc2;
        /// <summary>
        /// 二维数组
        /// </summary>
        public  List<List<int>> DoubleSlc1;
        /// <summary>
        /// 二维数组
        /// </summary>
        public  List<List<string>> DoubleSlc2;
        /// <summary>
        /// 布尔
        /// </summary>
        public bool IsBool;
        /// <summary>
        /// map类型
        /// </summary>
        public Dictionary<int,int> Map1;
        /// <summary>
        /// map类型
        /// </summary>
        public Dictionary<int,string> Map2;
    }

    public class AbbInfo
    {
        /// <summary>
        /// ID
        /// </summary>
        public int ID;
        /// <summary>
        /// 参数1
        /// </summary>
        public string Sth;
        /// <summary>
        /// 其他
        /// </summary>
        public string Other;
    }

    public class ExampleCfg
    {
        private static ExampleCfg _instance;
        public static ExampleCfg Get()
        {
            if (_instance == null)
            {
                _instance = Create();
            }
            return _instance;
        }
        public List<ExampleInfo> Example;
        private Dictionary<int, ExampleInfo> _ExampleDict;
        public List<AbbInfo> Abb;
        private Dictionary<int, AbbInfo> _AbbDict;
        private static ExampleCfg Create()
        {
            AssetHandle handle = YooAssets.LoadAssetSync<TextAsset>("Assets/GameRes/Json/Example");
            TextAsset text = handle.AssetObject as TextAsset;
            return JsonConvert.DeserializeObject<ExampleCfg>(text.text);
        }

        private void InitDict()
        {
            _ExampleDict = new Dictionary<int, ExampleInfo>();
            foreach(ExampleInfo info in Example)
            {
                _ExampleDict.Add(info.ID, info);
            }
            _AbbDict = new Dictionary<int, AbbInfo>();
            foreach(AbbInfo info in Abb)
            {
                _AbbDict.Add(info.ID, info);
            }
        }

        public ExampleInfo GetExampleByID(int id)
        {
            if (_ExampleDict == null)
            {
                InitDict();
            }
            return _ExampleDict[id];
        }
        public AbbInfo GetAbbByID(int id)
        {
            if (_AbbDict == null)
            {
                InitDict();
            }
            return _AbbDict[id];
        }
    }
}
