using System;
using System.Collections.Generic;
using System.Text;

namespace archiver
{
    public static class MyData
    {
        public static string a_出报告日期 = "2021年11月29日";
        public static string a_甲方单位 = "上海浦东发展银行股份有限公司";
        public static string a_系统名称 = "";
        public static int a_下标数量 = 3;
        public static List<string> a_xitong =new List<string>();
        public static string a_网络结构 = "";
        //原有的文字，责任划分、主管部门
        public static string a_被测对象描述_1 = "";
        // 网站的介绍，在2.1.2 外部网站系统是用以发布信息、树立企业形象、进……
        public static string a_被测对象描述_2 = "";
        //网站的防护，拓扑中有
        public static string a_被测对象描述_3 = "";
        //
        public static string a_安全状况描述_1;
        public static string a_安全状况描述_2 = ""; //问题发现和统计，在总体最后一段 (本次测评共发现安全问题29个....

        // 总体评价内容提取：
        public static string a_1总体评价_安全物理环境 = "";
        public static string a_2总体评价_安全通信网络 = "";
        public static string a_3总体评价_安全区域边界 = "";
        public static string a_4总体评价_安全计算环境方面 = "";
        public static string a_5总体评价_安全管理中心方面 = "";
        public static string a_6总体评价_安全管理制度方面 = "";
        public static string a_7总体评价_安全管理机构方面 = "";
        public static string a_8总体评价_安全管理人员方面 = "";
        public static string a_9总体评价_安全建设管理方面 = "";
        public static string a_10总体评价_安全运维管理 = "";

        // 4.3最后一列 严重程度变化
        public static string a_严重程度变化 = 
@"升高
√ 降低 ";


        // 1.1 测评依据
        public static string a_测评依据_金融 =
@"测评过程中主要依据的标准：
GB/T 22239-2019：《信息安全技术 网络安全等级保护基本要求》
GB/T 28448-2019：《信息安全技术 网络安全等级保护测评要求》
JR/T 0071.2-2020：《金融行业网络安全等级保护实施指引 第 2 部分：基本要求》
JR/T 0072—2020：《金融行业网络安全等级保护测评指南》
以下为本次测评的相关参考标准和文档：
GB/T 17859—1999 ：《计算机信息系统 安全保护等级划分准则》
GB/T 28449-2018：《信息安全技术 网络安全等级保护测评过程指南》
GB/T 20984-2007：《信息安全技术 信息安全风险评估规范》
";

// 1.1 测评依据
        public static string a_测评依据 = 
@"测评过程中主要依据的标准：
    GB/T 22239-2019：《信息安全技术 网络安全等级保护基本要求》
    GB/T 28448-2019：《信息安全技术 网络安全等级保护测评要求》
以下为本次测评的相关参考标准和文档：
    GB/T 17859—1999 ：《计算机信息系统 安全保护等级划分准则》
    GB/T 28449-2018：《信息安全技术 网络安全等级保护测评过程指南》
    GB/T 20984-2007：《信息安全技术 信息安全风险评估规范》
";

        // 1.4 等级测评报告正本一式二份，其中上海浦东发展银行股份有限公司一份，上海计算机软件技术开发中心一份。
        // 报告范围要换
        public static string a_报告范围 = @"等级测评报告正本一式二份，其中xxx公司一份，上海计算机软件技术开发中心一份。";
       
        
            //1.3测评过程
        public static string a_测评过程 = 
@"本次等级测评分为四个过程：测评准备过程、方案编制过程、测评实施过程、分析与报告编制过程。其中，各阶段的时间安排如下：
1、2021年07月12日～07月13日，测评准备过程。 
2、2021年07月14日～07月19日，方案编制过程。 
3、2021年07月20日～10月15日，现场实施过程。 
4、2021年10月18日～11月29日，分析与报告编制过程。 
其中，2021年07月12日召开了项目启动会议，确定了工作计划及项目人员名单；2021年10月15日召开了项目末次会议，确认了测评发现的问题。 
";

        //# 编号意义
        public static string a_编号意义 = @"第三组为机构代码，由网络安全等级测评与检测评估机构服务认证证书编号最后四位数字组成。";


    }
}
