using Cognex.VisionPro.Caliper;
using Cognex.VisionPro.ImageProcessing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YuanliCore.ImageProcess;
using YuanliCore.Interface;

namespace SPIL.model
{
    public class AlgorithmSetting: AbstractRecipe
    {
        /// <summary>
        /// 清晰度演算法
        /// </summary>
        public AlgorithmDescribe[] ClarityAlgorithms { get; set; }

        /// <summary>
        /// AOI 驗算法
        /// </summary>
        public AlgorithmDescribe[] AlgorithmDescribes { get; set; }

    }
    // 自訂的演算法顯示
    public class AlgorithmDescribe
    {
        public string Id { get; set; }
        public string Name { get; set; }

        public MethodType CogMethodtype  { get; set; }

       
        public CogMethod CogAOIMethod { get; set; }

        public AlgorithmDescribe(string id, string name, MethodType methodType)
        {
            Id = id;
            Name = name;
            CogMethodtype = methodType;

      //      CogSearchMaxTool;
      //      CogImageConvertTool asd;
    
        }

        public override string ToString()
        {
            return Id + " | " + Name;
        }
    }

    public enum MethodType
    {
        CogSearchMaxTool,
        CogImageConvertTool,
        CogFindEllipseTool,


    }
}
