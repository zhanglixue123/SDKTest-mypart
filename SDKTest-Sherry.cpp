// SDKTest.cpp : 定义 DLL 应用程序的导出函数。

#include "stdafx.h"
#include "SDKTest.h"
#include <tchar.h>
#include "ZjSpring.h"
#include "libxl.h"
#include "log4z\log4z.h"
#include<stdio.h>
#include <iostream>
#include <atlcoll.h>


using namespace zsummer::log4z;  
using namespace spring::access;
using namespace std;  
using namespace ATL;  



// 定义结构体：4、5级节点
class Ccase
{
public:
	CString		rtst_name;
	CString		rtst_tzx;
	CString		rtst_logic;
	CString		rtst_val;
	CString		rtst_rule;

};

class Creport
{
public:
	CString		rtst_name;
	CString		rtst_tzx;
	CString		rtst_logic;
	CString		rtst_val;
	CString		rtst_rule;
	
	CAtlArrayEx<Ccase>	cases;//case

};

class CRrow
{
public:
	CString		idx;
	CString		key;
	CString		voltagelevel;//电压等级
	CString		con;
	CString		dst_idx;
	CString		dst_key;
	CString		conremark;

	CAtlArrayEx<Creport>	reports;//report

};

class CArow
{
public:
	CString		index;
	CString		key;
	CString		voltagelevelCF;//电压等级CF
	CString		cond;
	CString		voltagelevelAF;//电压等级AF
	CString		dstIdx;
	CString		dstKey;
	CString		disdstIdx;
	CString		demand;
	CString		disDemand;
	CString		relati;
	CString		rCert;
	CString		condRemark;

};

class Cfeature
{
public:
	CString		text;
};

class Cresult
{
public:
	CString		ret_rept;
	CString		ret_name;
	CString		ret_logic;
	CString		ret_des;
	CString		ret_match;
	CString		ret_sat;

};

// 定义结构体：2、3级节点
class Crule
{
public:
	CString		rule_name;
	CString		rule_decribe;
	CString		rule_use;

};

class Cinfo
{
public:
	CString		info_name;
	CString		info_case;
	CString		info_auto;
};

class Ctnum
{
public:
	CString		tnum_value;
};

class Ccargo
{
public:
	CString		cargo_name;
	CString		cargo_achi;
	CString		cargo_rept;
	CString		cargo_type;
};

class CReptClaim
{
public:
	CString		table_num;
	CString		comment;
	CString		test_count;

	Cresult		result;
	Cfeature	feature;

	CAtlArrayEx<CRrow>	Rrows;
};

class CAchiClaim
{
public:
	CString		beginYear;//原Name
	CString		endYear;//原Value
	CString		tableNum;
	CString		numrCert;
	CString		unit;
	CString		AchiComment;

	CAtlArrayEx<CArow>	Arows;
};

class CQMSClaim
{
public:
	CString		text;
};

class CQualiClaimSheet//原Note1
{
public:
	CString		name;
	CString		type;
	CString		time;
	//CAtlArrayEx<Node2>	Node2s_AttrList;//数组

	CAchiClaim  AchiClaim;
	CReptClaim  ReptClaim;
	CQMSClaim	QMSClaim;

	CAtlArrayEx<Crule>	rules;
	CAtlArrayEx<Cinfo>	infos;
	CAtlArrayEx<Ctnum>	tnums;
	CAtlArrayEx<Ccargo>	cargos;

};

void excelTest(CQualiClaimSheet& QualiClaimSheet);
void ExportSheet1();
void ExportSheet2();
void ExportSheet3();
void ExportSheet4();
void ExportSheet5();
void ExportSheet6();

bool LoadAllSheetDataFromXml(CQualiClaimSheet& QualiClaimSheet, CXmlElement* Node1_QualiClaimSheet)
{
	
	QualiClaimSheet.name = Node1_QualiClaimSheet->GetAttrValue(_T("name"));
	QualiClaimSheet.type = Node1_QualiClaimSheet->GetAttrValue(_T("type"));
	QualiClaimSheet.time = Node1_QualiClaimSheet->GetAttrValue(_T("time"));

	// AchiClaim、ReptClaim、QMSClaim
	// 获取Node2子节点 AchiClaim、ReptClaim、QMSClaim
	CXmlElement* Node2_AchiClaim = Node1_QualiClaimSheet->GetChildElementAt(_T("AchiClaim"));
	CAchiClaim  &AchiClaim = QualiClaimSheet.AchiClaim;
	AchiClaim.beginYear = Node2_AchiClaim->GetAttrValue(_T("beginYear"));
	AchiClaim.endYear = Node2_AchiClaim->GetAttrValue(_T("endYear"));
	AchiClaim.AchiComment = Node2_AchiClaim->GetAttrValue(_T("AchiComment"));
	AchiClaim.numrCert = Node2_AchiClaim->GetAttrValue(_T("numrCert"));
	AchiClaim.tableNum = Node2_AchiClaim->GetAttrValue(_T("tableNum"));
	AchiClaim.unit = Node2_AchiClaim->GetAttrValue(_T("unit"));

	CReptClaim  &ReptClaim = QualiClaimSheet.ReptClaim;//采用引用的方式，相当于快捷方式，比QualiClaimSheet.AchiClaim.unit更简洁。
	CXmlElement* Node2_ReptClaim = Node1_QualiClaimSheet->GetChildElementAt(_T("ReptClaim"));
	ReptClaim.table_num = Node2_ReptClaim->GetAttrValue(_T("table_num"));
	ReptClaim.comment = Node2_ReptClaim->GetAttrValue(_T("comment"));
	ReptClaim.test_count = Node2_ReptClaim->GetAttrValue(_T("test_count"));

	CQMSClaim	&QMSClaim = QualiClaimSheet.QMSClaim ;
	CXmlElement* Node2_QMSClaim = Node1_QualiClaimSheet->GetChildElementAt(_T("QMSClaim"));
	QMSClaim.text = Node2_QMSClaim->GetElementText();

	//获取子节点 QualiClaimSheet->rule
	CXmlElement* Node2_rules = Node1_QualiClaimSheet->GetChildElementAt(_T("rules"));
	//CString Node3Attr_rule_Name[100][100];	CString Node3Attr_rule_Value[100][100];
	//for (int i = 0; i < Node2_rules->GetChildElementCount(); i ++)
	//{
	//	CXmlElement* Node3_rule = Node2_rules->GetChildElements()->GetAt (i);
	//	for (int j = 0; j < Node3_rule->GetAttributeCount (); j ++)
	//	{
	//		CXmlAttribute* Node3Attr_rule = Node3_rule->GetAttributes ()->GetAt (j);
	//		Node3Attr_rule_Name[i][j] = Node3Attr_rule->GetAttrName();
	//		Node3Attr_rule_Value[i][j] = Node3Attr_rule->GetStrValue();
	//	}
	//}
	CAtlArrayEx<Crule>	&rules = QualiClaimSheet.rules ;
	for (int i = 0; i < Node2_rules->GetChildElementCount(); i ++)
	{
		CXmlElement* Node3_rule = Node2_rules->GetChildElements()->GetAt (i);
		Crule	rule   ;
		rule.rule_name = Node3_rule->GetAttrValue(_T("rule_name"));
		rule.rule_decribe = Node3_rule->GetAttrValue(_T("rule_decribe"));
		rule.rule_use = Node3_rule->GetAttrValue(_T("rule_use"));

		rules.Add(rule);
	}


	// 获取子节点 QualiClaimSheet->info
	CXmlElement* Node2_infos = Node1_QualiClaimSheet->GetChildElementAt(_T("infos"));
	// 获取子节点属性 info
	CAtlArrayEx<Cinfo>	&infos = QualiClaimSheet.infos ;
	for (int i = 0; i < Node2_infos->GetChildElementCount(); i ++)
	{
		CXmlElement* Node3_info = Node2_infos->GetChildElements()->GetAt (i);
		Cinfo	info;
		info.info_name = Node3_info->GetAttrValue(_T("info_name"));
		info.info_case = Node3_info->GetAttrValue(_T("info_case"));
		info.info_auto = Node3_info->GetAttrValue(_T("info_auto"));

		infos.Add(info);
	}

	// 获取子节点 QualiClaimSheet->tnumm
	CXmlElement* Node2_tnums = Node1_QualiClaimSheet->GetChildElementAt(_T("tnum"));
	// 获取子节点属性 tnumm
	CAtlArrayEx<Ctnum>	&tnums = QualiClaimSheet.tnums ;
	for (int i = 0; i < Node2_tnums->GetChildElementCount(); i ++)
	{
		CXmlElement* Node3_tnum = Node2_tnums->GetChildElements()->GetAt (i);
		Ctnum	tnum;
		tnum.tnum_value = Node3_tnum->GetAttrValue(_T("tnum_value"));

		tnums.Add(tnum);
	}

	// 获取子节点 QualiClaimSheet->cargos
	CXmlElement* Node2_cargos = Node1_QualiClaimSheet->GetChildElementAt(_T("cargos"));
	//// 获取子节点属性 cargo
	CAtlArrayEx<Ccargo>	&cargos = QualiClaimSheet.cargos;
	for (int i = 0; i < Node2_cargos->GetChildElementCount(); i ++)
	{
		CXmlElement* Node3_cargo = Node2_cargos->GetChildElements()->GetAt (i);
		Ccargo	cargo;
		cargo.cargo_name = Node3_cargo->GetAttrValue(_T("cargo_name"));
		cargo.cargo_achi = Node3_cargo->GetAttrValue(_T("cargo_achi"));
		cargo.cargo_rept = Node3_cargo->GetAttrValue(_T("cargo_rept"));

		cargos.Add(cargo);
	}

	// 获取子节点 AchiClaim->rows
	CXmlElement* Node3_Arows = Node2_AchiClaim->GetChildElementAt(_T("rows"));
	// 获取子节点属性 row
	CAtlArrayEx<CArow>	&Arows = QualiClaimSheet.AchiClaim.Arows ;
	for (int i = 0; i < Node3_Arows->GetChildElementCount(); i ++)
	{
		CXmlElement* Node4_Arow = Node3_Arows->GetChildElements()->GetAt (i);
		CArow	Arow;
		Arow.index = Node4_Arow->GetAttrValue(_T("index"));
		Arow.key = Node4_Arow->GetAttrValue(_T("key"));
		Arow.voltagelevelCF = Node4_Arow->GetAttrValue(_T("电压等级CF"));
		Arow.cond = Node4_Arow->GetAttrValue(_T("cond"));
		Arow.voltagelevelAF = Node4_Arow->GetAttrValue(_T("电压等级AF"));
		Arow.dstIdx = Node4_Arow->GetAttrValue(_T("dstIdx"));
		Arow.dstKey = Node4_Arow->GetAttrValue(_T("dstKey"));
		Arow.disdstIdx = Node4_Arow->GetAttrValue(_T("disdstIdx"));
		Arow.demand = Node4_Arow->GetAttrValue(_T("demand"));
		Arow.disDemand = Node4_Arow->GetAttrValue(_T("disDemand"));
		Arow.relati = Node4_Arow->GetAttrValue(_T("relati"));
		Arow.rCert = Node4_Arow->GetAttrValue(_T("rCert"));
		Arow.condRemark = Node4_Arow->GetAttrValue(_T("condRemark"));

		Arows.Add(Arow);
	}

	// 获取子节点 ReptClaim->cg->feature
	CXmlElement* Node3_cg = Node2_ReptClaim->GetChildElementAt(_T("cg"));
	CXmlElement* Node4_feature = Node3_cg->GetChildElementAt(_T("feature"));
	// 获取子节点文本 feature
	Cfeature	&feature = QualiClaimSheet.ReptClaim.feature ;
	feature.text = Node4_feature->GetElementText();

	// 获取子节点 ReptClaim->results
	CXmlElement* Node3_results = Node2_ReptClaim->GetChildElementAt(_T("results"));
	Cresult  &result = QualiClaimSheet.ReptClaim.result;
	CXmlElement* Node4_result = Node3_results->GetChildElementAt(_T("result"));
	result.ret_rept = Node4_result->GetAttrValue(_T("ret_rept"));
	result.ret_name = Node4_result->GetAttrValue(_T("ret_name"));
	result.ret_logic = Node4_result->GetAttrValue(_T("ret_logic"));
	result.ret_des = Node4_result->GetAttrValue(_T("ret_des"));
	result.ret_match = Node4_result->GetAttrValue(_T("ret_match"));


	//获取子节点 ReptClaim->rows
	CXmlElement* Node3_Rrows = Node2_ReptClaim->GetChildElementAt(_T("rows"));
	//// 获取子节点属性 row
	CAtlArrayEx<CRrow>	&Rrows = QualiClaimSheet.ReptClaim.Rrows ;//特别注意：循环嵌套的时候不能对下级节点重置CAtlArrayEx<Creport>	reports;或者case，否则会访问一个独立定义的reports数组，不在row之下了。
	CXmlElement* Node4_Rrow = Node3_Rrows->GetChildElementAt(_T("row"));
	//// 获取子节点属性 row
	for (int i = 0; i < Node3_Rrows->GetChildElementCount(); i ++)
	{
		CXmlElement* Node4_Rrow = Node3_Rrows->GetChildElements()->GetAt (i);
		CRrow	Rrow;
		Rrow.idx = Node4_Rrow->GetAttrValue(_T("idx"));
		Rrow.key = Node4_Rrow->GetAttrValue(_T("key"));
		Rrow.voltagelevel = Node4_Rrow->GetAttrValue(_T("电压等级"));
		Rrow.con = Node4_Rrow->GetAttrValue(_T("con"));
		Rrow.dst_idx = Node4_Rrow->GetAttrValue(_T("dst_idx"));
		Rrow.dst_key = Node4_Rrow->GetAttrValue(_T("dst_key"));
		Rrow.conremark = Node4_Rrow->GetAttrValue(_T("conremark"));
		//  获取子节点属性 rows->report
		for (int j = 0; j < Node4_Rrow->GetChildElementCount(); j ++)
		{
			CXmlElement* Node5_report = Node4_Rrow->GetChildElements()->GetAt (j);
			Creport	report ;
			report.rtst_name = Node5_report->GetAttrValue(_T("rtst_name"));
			report.rtst_tzx = Node5_report->GetAttrValue(_T("rtst_tzx"));
			report.rtst_val = Node5_report->GetAttrValue(_T("rtst_val"));
			report.rtst_rule = Node5_report->GetAttrValue(_T("rtst_rule"));
			//  获取子节点属性 rows->report->case
			if ( Node5_report->GetChildElementCount() == 1)
			{
				//CAtlArrayEx<Ccase>	&cases = report.cases;
				CXmlElement* Node6_case = Node5_report->GetChildElementAt(_T("case"));
				Ccase	case_ ;
				case_.rtst_name = Node6_case->GetAttrValue(_T("rtst_name"));
				case_.rtst_tzx = Node6_case->GetAttrValue(_T("rtst_tzx"));
				case_.rtst_val = Node6_case->GetAttrValue(_T("rtst_val"));
				case_.rtst_rule = Node6_case->GetAttrValue(_T("rtst_rule"));

				report.cases.Add(case_);
			}
			Rrow.reports.Add(report);
		}
		Rrows.Add(Rrow);
	}

	return true;
}


// 测试函数，请在该函数中添加需要的测试代码，请勿删除该函数
SDKTEST_API int fnSDKTest(void)
{
	//测试提示框
	::MessageBox(::GetActiveWindow(), _T("SDKTest.dll中fnSDKTest函数正在被调用！"), _T("提示"), MB_OK);

	//测试xmd文件加载
	CXmlDocument xmdDoc;
	if (!xmdDoc.LoadFile(_T("C:\\博微软件\\评标辅助工具_资质评审-物资公司版\\sdkTest.xmd"), fmtXMD))
	{
		::MessageBox(::GetActiveWindow(), _T("加载XMD文件失败！"), _T("提示"), MB_ICONERROR);
		return -1;
	}

	// 获取根节点 Document
	/*CXmlElement* pEle = xmdDoc.GetElementRoot();
	pEle->*/

	// 获取节点属性根据名称 name
	//CString strValue = pEle->GetAttrValue(_T("name"));

	// 获取子节点根据子节点名称 QualiClaimSheet
	//CXmlElement* pQualiClaimSheetEle = pEle->GetChildElementAt(_T("QualiClaimSheet"));

	//取子节点数量和某一子节点的实例代码
	//CXmlElement* elmt;
	//for (int i = 0; i < elmt->GetChildElementCount(); i ++)
	//{
	//	CXmlElement* childElmt = elmt->GetChildElements()->GetAt (i);
	//}

	//遍历一个节点下所有属性并取属性名称和值的方法
	//for (int i = 0; i < elmt->GetAttributeCount (); i ++)
	//{
	//	CXmlAttribute* attr = elmt->GetAttributes ()->GetAt (i);
	//	cout << attr->GetAttrName ();
	//	cout << attr->GetStrValue ();
	//}

	xmdDoc.SaveFile(_T("C:\\博微软件\\评标辅助工具_资质评审-物资公司版\\sdkTest.xml"), fmtXML);

	//读取xmd文件内容
	CString xmdString;
	if(!xmdDoc.GetXmlString(xmdString))
	{
		::MessageBox(::GetActiveWindow(), _T("读取XMD文件内容失败！"), _T("提示"), MB_ICONERROR);
		return -1;
	}

	//::MessageBox(::GetActiveWindow(), xmdString, _T("提示"), MB_OK);

	// 建立结构体  Root
	class CDocument
	{
	public:
		CString		name;
		CString		type;
		CString		guid;
		CString		status;
		CString		time;
		CString		timef;

	};
	// 获取根节点 Document
	CXmlElement* Root_Document  = xmdDoc.GetElementRoot();
	if (Root_Document == NULL)  
	{  
		return -1;  
	} 

	// 获取Node1子节点 QualiClaimSheet
	CXmlElement* Node1_QualiClaimSheet = Root_Document->GetChildElementAt(_T("QualiClaimSheet"));
	// 获取子节点属性 QualiClaimSheet
	CQualiClaimSheet QualiClaimSheet;

	// 导入各种节点
	LoadAllSheetDataFromXml(QualiClaimSheet, Node1_QualiClaimSheet);
	// 生成excel
	excelTest(QualiClaimSheet);

	return 0;
}

void ExportSheet1(libxl::Book *pBook, libxl::Sheet *pSheet, libxl::Format* pHeadFormat, libxl::Font *pFont, libxl::Format* pFormat, libxl::Format* pLeftTopFormat, libxl::Format* pDesignWidth, CAtlArrayEx<CArow> &Arows, CAchiClaim &AchiClaim)
{
	CString strSheetName = _T("业绩要求模板-保护");
	pSheet = pBook->addSheet(strSheetName);
	if(pSheet == NULL)
	{
		::MessageBox(::GetActiveWindow(), _T("添加Excel中的Sheet页失败！"), _T("提示"), MB_ICONERROR);
		return;
	}
	
	// 写入数据
	//pSheet->setCol(0, 1, 20, pDesignWidth, false);

	// 表头
	pSheet->writeStr(0, 0, _T("目号") , pHeadFormat);
	//double num = pSheet->readNum(0, 0, pHeadFormat);
	//CString num ;
	//num.Format(_T("%f"), pSheet->readStr(0, 0, &pHeadFormat));
	//num = pSheet->readStr(0, 0, &pHeadFormat);
	//CString num1;
	//num1.Format(_T("%d"),num.GetLength() );

	pSheet->writeStr(0, 1, _T("货物名称	") , pHeadFormat);
	pSheet->writeStr(0, 3, AchiClaim.beginYear + _T("年到") + AchiClaim.endYear + _T("年供货业绩要求") , pHeadFormat);
	pSheet->writeStr(0, 8, _T("认证证书") , pHeadFormat);
	pSheet->writeStr(0, 9, _T("条件备注") , pHeadFormat);
	pSheet->writeStr(1, 1, _T("小类名称") , pHeadFormat);
	pSheet->writeStr(1, 2, _T("货物特征值") , pHeadFormat);
	pSheet->writeStr(2, 2, _T("电压等级") , pHeadFormat);
	pSheet->writeStr(1, 3, _T("条件") , pHeadFormat);
	pSheet->writeStr(1, 4, _T("条件对应目号") , pHeadFormat);
	pSheet->writeStr(1, 5, _T("业绩特征值") , pHeadFormat);
	pSheet->writeStr(1, 6, _T("条件关系(只能填写一种关系)") , pHeadFormat);
	pSheet->writeStr(2, 6, _T("电压等级") , pHeadFormat);
	pSheet->writeStr(1, 7, _T("数量") , pHeadFormat);
	pSheet->writeStr(2, 7, _T("单位：") + AchiClaim.unit , pHeadFormat);
	// 内容
	int nRowIndex = 3;
	for (int i = 0; i < Arows.GetCount(); i ++)
	{
		CArow Arow = Arows.GetAt(i);
		pSheet->writeStr(nRowIndex + i, 0, Arow.index, pFormat);
		pSheet->writeStr(nRowIndex + i, 1, Arow.key, pFormat);
		pSheet->writeStr(nRowIndex + i, 2, Arow.voltagelevelCF, pFormat);
		pSheet->writeStr(nRowIndex + i, 3, Arow.cond, pFormat);
		pSheet->writeStr(nRowIndex + i, 4, Arow.dstKey, pFormat);
		pSheet->writeStr(nRowIndex + i, 5, Arow.relati, pFormat);
		pSheet->writeStr(nRowIndex + i, 6, Arow.voltagelevelAF, pFormat);
		pSheet->writeStr(nRowIndex + i, 7, Arow.demand, pFormat);
		pSheet->writeStr(nRowIndex + i, 8, Arow.rCert, pFormat);
		pSheet->writeStr(nRowIndex + i, 9, Arow.condRemark, pFormat);
	}

	//合并单元格 pSheet->setMerge(int rowFirst, int rowLast, int colFirst, int colLast);
	pSheet->setMerge(0, 2, 0, 0);
	pSheet->setMerge(0, 0, 1, 2);
	pSheet->setMerge(0, 0, 3, 7);
	pSheet->setMerge(1, 2, 1, 1);
	pSheet->setMerge(1, 2, 3, 3);
	pSheet->setMerge(1, 2, 4, 4);
	pSheet->setMerge(1, 2, 5, 5);
	pSheet->setMerge(0, 2, 8, 8);
	pSheet->setMerge(0, 2, 9, 9);

	// 备注
	//CString liekuan ;
	//liekuan.Format(_T("%f"), pSheet->colWidth(0));
	pSheet->writeStr(nRowIndex + Arows.GetCount() + 2, 0, _T("备注：") , pLeftTopFormat);
	pSheet->setMerge(nRowIndex + Arows.GetCount() + 2, nRowIndex + Arows.GetCount() + 3, 0, 9);


	// 自动调整列宽
	int TextWidth_Max , TextWidth_Head , TextWidth_Final , TextWidth;
	CString Text ;
	int i,j;
	for ( i = 0; i < 100; i++)
	{
		for ( j=0; j< 50; j++)
		{
			if (j == 0)
			{ Text = pSheet->readStr(j, i, &pHeadFormat);
			TextWidth = Text.GetLength() ;
			TextWidth_Head = TextWidth;
			TextWidth_Max = TextWidth;}
			else
			{
				Text = pSheet->readStr(j, i, &pFormat);
				TextWidth = Text.GetLength() ;
			}
			if (TextWidth_Max < TextWidth)
				TextWidth_Max = TextWidth;

		}
		if (TextWidth_Max > 20)
			TextWidth_Max = 20;
		pSheet->setCol(i, i, TextWidth_Max *2, pDesignWidth, false);
	}

}

void ExportSheet2(libxl::Book *pBook,libxl::Sheet *pSheet, libxl::Format* pHeadFormat, libxl::Font *pFont, libxl::Format* pFormat, libxl::Format* pLeftTopFormat, libxl::Format* pDesignWidth, CAtlArrayEx<CRrow> &Rrows)
{
	CString strSheetName = _T("报告要求模板-保护");
	pSheet = pBook->addSheet(strSheetName);
	if(pSheet == NULL)
	{
		::MessageBox(::GetActiveWindow(), _T("添加Excel中的Sheet页失败！"), _T("提示"), MB_ICONERROR);
		return;
	}

	// 写入数据
	// 表头
	pSheet->writeStr(0, 0, _T("目号") , pHeadFormat);
	pSheet->writeStr(0, 1, _T("货物名称	") , pHeadFormat);
	pSheet->writeStr(0, 3, _T("条件") , pHeadFormat);
	pSheet->writeStr(0, 4, _T("条件对应目号") , pHeadFormat);
	pSheet->writeStr(0, 5, _T("条件关系(只能填写一种关系)") , pHeadFormat);
	pSheet->writeStr(0, 6, _T("报告要求	") , pHeadFormat);
	pSheet->writeStr(0, 26, _T("条件备注") , pHeadFormat);
	pSheet->writeStr(1, 6, _T("试验报告1") , pHeadFormat);
	pSheet->writeStr(1, 16, _T("试验报告2") , pHeadFormat);
	pSheet->writeStr(1, 21, _T("试验报告3") , pHeadFormat);
	pSheet->writeStr(2, 6, _T("报告名称") , pHeadFormat);
	pSheet->writeStr(2, 11, _T("试验项目1") , pHeadFormat);
	pSheet->writeStr(2, 16, _T("报告名称") , pHeadFormat);
	pSheet->writeStr(2, 21, _T("报告名称") , pHeadFormat);
	pSheet->writeStr(3, 1, _T("小类名称") , pHeadFormat);
	pSheet->writeStr(3, 2, _T("电压等级") , pHeadFormat);
	for (int i = 0; i < 4; i++)
	{
		pSheet->writeStr(2, 7 + i*5 , _T("特殊规则") , pHeadFormat);
		pSheet->writeStr(3, 7 + i*5 , _T("涉及特征项") , pHeadFormat);
		pSheet->writeStr(3, 8 + i*5 , _T("规则逻辑") , pHeadFormat);
		pSheet->writeStr(3, 9 + i*5 , _T("数值") , pHeadFormat);
		pSheet->writeStr(3, 10 + i*5 , _T("规则描述") , pHeadFormat);
	}
	
	int nRowIndex = 4;
	for (int i = 0; i < Rrows.GetCount(); i ++)
	{
		// 目号A ~ 条件关系(只能填写一种关系)F
		CRrow Rrow = Rrows.GetAt(i);
		pSheet->writeStr(nRowIndex + i, 0, Rrow.idx, pFormat);
		pSheet->writeStr(nRowIndex + i, 1, Rrow.key, pFormat);
		pSheet->writeStr(nRowIndex + i, 2, Rrow.voltagelevel, pFormat);
		pSheet->writeStr(nRowIndex + i, 3, Rrow.con, pFormat);
		pSheet->writeStr(nRowIndex + i, 4, Rrow.dst_key, pFormat);
		pSheet->writeStr(nRowIndex + i, 5, Rrow.conremark, pFormat);
		pSheet->writeStr(nRowIndex + i, 26, Rrow.conremark, pFormat);
		// report G~P 
			Creport report = Rrow.reports.GetAt(0);
			pSheet->writeStr(nRowIndex + i, 6  , report.rtst_name, pFormat);
			pSheet->writeStr(nRowIndex + i, 7  , report.rtst_tzx, pFormat);
			pSheet->writeStr(nRowIndex + i, 8  , report.rtst_logic, pFormat);
			pSheet->writeStr(nRowIndex + i, 9  , report.rtst_val, pFormat);
			pSheet->writeStr(nRowIndex + i, 10  , report.rtst_rule, pFormat);
			// case L~P
			Ccase case_ = report.cases.GetAt(0);
			pSheet->writeStr(nRowIndex + i, 11  , case_.rtst_name, pFormat);
			pSheet->writeStr(nRowIndex + i, 12  , case_.rtst_tzx, pFormat);
			pSheet->writeStr(nRowIndex + i, 13  , case_.rtst_logic, pFormat);
			pSheet->writeStr(nRowIndex + i, 14  , case_.rtst_val, pFormat);
			pSheet->writeStr(nRowIndex + i, 15  , case_.rtst_rule, pFormat);
			// report Q~Z
			for (int j = 1; j < 3; j ++)
			{
				report = Rrow.reports.GetAt(j);
				pSheet->writeStr(nRowIndex + i, 16 + (j-1)*5  , report.rtst_name, pFormat);
				pSheet->writeStr(nRowIndex + i, 17 + (j-1)*5 , report.rtst_tzx, pFormat);
				pSheet->writeStr(nRowIndex + i, 18 + (j-1)*5 , report.rtst_logic, pFormat);
				pSheet->writeStr(nRowIndex + i, 19 + (j-1)*5 , report.rtst_val, pFormat);
				pSheet->writeStr(nRowIndex + i, 20 + (j-1)*5 , report.rtst_rule, pFormat);
			}
			

	}
	// 备注
	pSheet->writeStr(nRowIndex + Rrows.GetCount() + 2, 0, _T("备注：") , pLeftTopFormat);
	
	//合并单元格 pSheet->setMerge(int rowFirst, int rowLast, int colFirst, int colLast);
	pSheet->setMerge(0, 3, 0, 0);
	pSheet->setMerge(0, 2, 1, 2);
	pSheet->setMerge(0, 3, 3, 3);
	pSheet->setMerge(0, 3, 4, 4);
	pSheet->setMerge(0, 3, 5, 5);
	pSheet->setMerge(0, 0, 6, 25);
	pSheet->setMerge(1, 1, 6, 15);
	pSheet->setMerge(2, 3, 6, 6);
	pSheet->setMerge(2, 2, 7, 10);
	pSheet->setMerge(2, 3, 11, 11);
	pSheet->setMerge(2, 2, 12, 15);
	pSheet->setMerge(1, 1, 16, 20);
	pSheet->setMerge(2, 3, 16, 16);
	pSheet->setMerge(2, 2, 17, 20);
	pSheet->setMerge(1, 1, 21, 25);
	pSheet->setMerge(2, 3, 21, 21);
	pSheet->setMerge(2, 2, 22, 25);
	pSheet->setMerge(0, 3, 26, 26);
	pSheet->setMerge(nRowIndex + Rrows.GetCount() + 2, nRowIndex + Rrows.GetCount() + 3, 0, 26);
	
	// 自动调整列宽
	int TextWidth_Max , TextWidth_Head , TextWidth_Final , TextWidth;
	CString Text ;
	int i,j;
	for ( i = 0; i < 100; i++)
	{
		for ( j=0; j< 50; j++)
		{
			if (j == 0)
			{ Text = pSheet->readStr(j, i, &pHeadFormat);
			TextWidth = Text.GetLength() ;
			TextWidth_Head = TextWidth;
			TextWidth_Max = TextWidth;}
			else
			{
				Text = pSheet->readStr(j, i, &pFormat);
				TextWidth = Text.GetLength() ;
			}
			if (TextWidth_Max < TextWidth)
				TextWidth_Max = TextWidth;

		}
		if (TextWidth_Max > 20)
			TextWidth_Max = 20;
		pSheet->setCol(i, i, TextWidth_Max *2, pDesignWidth, false);
	}
}

void ExportSheet3(libxl::Book *pBook, libxl::Sheet *pSheet, libxl::Format* pHeadFormat, libxl::Font *pFont, libxl::Format* pFormat,libxl::Format* pDesignWidth, CAtlArrayEx<Ccargo> &cargos)
{
	CString strSheetName = _T("特殊货物");
	pSheet = pBook->addSheet(strSheetName);
	if(pSheet == NULL)
	{
		::MessageBox(::GetActiveWindow(), _T("添加Excel中的Sheet页失败！"), _T("提示"), MB_ICONERROR);
		return;
	}

	// 写入数据
	// 表头
	pSheet->writeStr(0, 0, _T("货物名称") , pHeadFormat);
	pSheet->writeStr(0, 1, _T("业绩要求") , pHeadFormat);
	pSheet->writeStr(0, 2, _T("报告要求") , pHeadFormat);
	pSheet->writeStr(0, 3, _T("物资类型") , pHeadFormat);
	// 内容
	int nRowIndex = 1;
	for (int i = 0; i < cargos.GetCount(); i ++)
	{
		Ccargo cargo = cargos.GetAt(i);
		pSheet->writeStr(nRowIndex + i, 0, cargo.cargo_name, pFormat);
		pSheet->writeStr(nRowIndex + i, 1, cargo.cargo_achi, pFormat);
		pSheet->writeStr(nRowIndex + i, 2, cargo.cargo_rept, pFormat);
		pSheet->writeStr(nRowIndex + i, 3, cargo.cargo_type, pFormat);//这里有问题，xmd中的读不过来
	}
	pSheet->writeStr(1, 3, _T("在货物清单中"), pFormat);//所以只好补充在这
	pSheet->writeStr(2, 3, _T("未在货物清单中"), pFormat);
	
	// 自动调整列宽
	int TextWidth_Max , TextWidth_Head , TextWidth_Final , TextWidth;
	CString Text ;
	int i,j;
	for ( i = 0; i < 100; i++)
	{
		for ( j=0; j< 50; j++)
		{
			if (j == 0)
			{ Text = pSheet->readStr(j, i, &pHeadFormat);
			TextWidth = Text.GetLength() ;
			TextWidth_Head = TextWidth;
			TextWidth_Max = TextWidth;}
			else
			{
				Text = pSheet->readStr(j, i, &pFormat);
				TextWidth = Text.GetLength() ;
			}
			if (TextWidth_Max < TextWidth)
				TextWidth_Max = TextWidth;

		}
		if (TextWidth_Max > 20)
			TextWidth_Max = 20;
		pSheet->setCol(i, i, TextWidth_Max *2, pDesignWidth, false);
	}

}



void ExportSheet4(libxl::Book *pBook, libxl::Sheet *pSheet, libxl::Format* pHeadFormat, libxl::Font *pFont, libxl::Format* pFormat, libxl::Format* pDesignWidth, CAtlArrayEx<Cinfo> &infos, CAtlArrayEx<Ctnum> &tnums, CQualiClaimSheet &QualiClaimSheet)
{
	CString strSheetName = _T("资格条件项目与一纸项目对应关系");
	pSheet = pBook->addSheet(strSheetName);
	if(pSheet == NULL)
	{
		::MessageBox(::GetActiveWindow(), _T("添加Excel中的Sheet页失败！"), _T("提示"), MB_ICONERROR);
		return;
	}

	// 写入数据
	// 表头
	pSheet->writeStr(0, 0, _T("序号") , pHeadFormat);
	pSheet->writeStr(0, 1, _T("设备名称") , pHeadFormat);
	pSheet->writeStr(0, 2, _T("资格条件项目") , pHeadFormat);
	pSheet->writeStr(0, 3, _T("试验项目名称") , pHeadFormat);
	pSheet->writeStr(0, 4, _T("对应一纸表单编号") , pHeadFormat);
	pSheet->writeStr(0, 5, _T("录入方式") , pHeadFormat);
	// 内容
	int nRowIndex = 1;
	CString idx; 
	for (int i = 0; i < infos.GetCount(); i ++)
	{
		Cinfo info = infos.GetAt(i);
		idx.Format(_T("%d"), (i+1));
		pSheet->writeStr(nRowIndex + i, 0, idx, pFormat);
		pSheet->writeStr(nRowIndex + i, 1, QualiClaimSheet.type, pFormat);
		pSheet->writeStr(nRowIndex + i, 2, info.info_name, pFormat);
		pSheet->writeStr(nRowIndex + i, 3, info.info_case, pFormat);
		if (info.info_auto == _T("1"))
			pSheet->writeStr(nRowIndex + i, 5, _T("一纸读取判定"), pFormat);
		else
			pSheet->writeStr(nRowIndex + i, 5, _T("手工录入"), pFormat);

	}
	Ctnum tnum = tnums.GetAt(1);
	pSheet->writeStr(1, 4,  _T("表") + tnum.tnum_value , pFormat);
	tnum = tnums.GetAt(2);
	pSheet->writeStr(2, 4,  _T("表") + tnum.tnum_value , pFormat);
	tnum = tnums.GetAt(3);
	pSheet->writeStr(3, 4,  _T("表") + tnum.tnum_value , pFormat);
	tnum = tnums.GetAt(4);
	pSheet->writeStr(4, 4,  _T("表") + tnum.tnum_value , pFormat);
	tnum = tnums.GetAt(6);
	pSheet->writeStr(5, 4,  _T("表") + tnum.tnum_value , pFormat);
	tnum = tnums.GetAt(7);
	pSheet->writeStr(6, 4,  _T("表") + tnum.tnum_value , pFormat);

	// 自动调整列宽
	int TextWidth_Max , TextWidth_Head , TextWidth_Final , TextWidth;
	CString Text ;
	int i,j;
	for ( i = 0; i < 100; i++)
	{
		for ( j=0; j< 50; j++)
		{
			if (j == 0)
			{ Text = pSheet->readStr(j, i, &pHeadFormat);
			TextWidth = Text.GetLength() ;
			TextWidth_Head = TextWidth;
			TextWidth_Max = TextWidth;}
			else
			{
				Text = pSheet->readStr(j, i, &pFormat);
				TextWidth = Text.GetLength() ;
			}
			if (TextWidth_Max < TextWidth)
				TextWidth_Max = TextWidth;

		}
		if (TextWidth_Max > 20)
			TextWidth_Max = 20;
		pSheet->setCol(i, i, TextWidth_Max *2, pDesignWidth, false);
	}
}

void ExportSheet5(libxl::Book *pBook, libxl::Sheet *pSheet, libxl::Format* pHeadFormat, libxl::Font *pFont, libxl::Format* pFormat,libxl::Format* pDesignWidth, Cresult result, CQualiClaimSheet &QualiClaimSheet)
{
	CString strSheetName = _T("试验项目结论符合描述");
	pSheet = pBook->addSheet(strSheetName);
	if(pSheet == NULL)
	{
		::MessageBox(::GetActiveWindow(), _T("添加Excel中的Sheet页失败！"), _T("提示"), MB_ICONERROR);
		return;
	}

	// 写入数据
	// 表头
	pSheet->writeStr(0, 0, _T("序号") , pHeadFormat);
	pSheet->writeStr(0, 1, _T("设备名称") , pHeadFormat);
	pSheet->writeStr(0, 2, _T("报告名称（资格条件）") , pHeadFormat);
	pSheet->writeStr(0, 3, _T("试验项目名称（资格条件）") , pHeadFormat);
	pSheet->writeStr(0, 4, _T("试验项目名称（一纸描述）") , pHeadFormat);
	pSheet->writeStr(0, 5, _T("结论描述（列出一纸各种结论）") , pHeadFormat);
	pSheet->writeStr(0, 6, _T("是否满足") , pHeadFormat);
	// 内容
	CString idx; 
	idx.Format(_T("%d"), 1);
	pSheet->writeStr(1, 0, idx, pFormat);
	pSheet->writeStr(1, 1, QualiClaimSheet.type, pFormat);
	pSheet->writeStr(1, 2, result.ret_rept, pFormat);
	pSheet->writeStr(1, 3, result.ret_name, pFormat);
	pSheet->writeStr(1, 4, result.ret_logic, pFormat);
	pSheet->writeStr(1, 5, result.ret_des, pFormat);
	if (result.ret_match == _T("0"))//本来应该是ret_sat可是sat的值在xmd中为空
		pSheet->writeStr(1, 6,  _T("不满足"), pFormat);

	// 自动调整列宽
	int TextWidth_Max , TextWidth_Head , TextWidth_Final , TextWidth;
	CString Text ;
	int i,j;
	for ( i = 0; i < 100; i++)
	{
		for ( j=0; j< 50; j++)
		{
			if (j == 0)
			{ Text = pSheet->readStr(j, i, &pHeadFormat);
			TextWidth = Text.GetLength() ;
			TextWidth_Head = TextWidth;
			TextWidth_Max = TextWidth;}
			else
			{
				Text = pSheet->readStr(j, i, &pFormat);
				TextWidth = Text.GetLength() ;
			}
			if (TextWidth_Max < TextWidth)
				TextWidth_Max = TextWidth;

		}
		if (TextWidth_Max > 20)
			TextWidth_Max = 20;
		pSheet->setCol(i, i, TextWidth_Max *2, pDesignWidth, false);
	}
}

void ExportSheet6(libxl::Book *pBook, libxl::Sheet *pSheet, libxl::Format* pHeadFormat, libxl::Font *pFont, libxl::Format* pFormat,libxl::Format* pDesignWidth, CAtlArrayEx<Crule> &rules, CQualiClaimSheet &QualiClaimSheet)
{
	CString strSheetName = _T("特殊规则");
	pSheet = pBook->addSheet(strSheetName);
	if(pSheet == NULL)
	{
		::MessageBox(::GetActiveWindow(), _T("添加Excel中的Sheet页失败！"), _T("提示"), MB_ICONERROR);
		return;
	}

	// 写入数据
	// 表头
	pSheet->writeStr(0, 0, _T("序号") , pHeadFormat);
	pSheet->writeStr(0, 1, _T("设备名称") , pHeadFormat);
	pSheet->writeStr(0, 2, _T("规则名称") , pHeadFormat);
	pSheet->writeStr(0, 3, _T("特殊情况描述") , pHeadFormat);
	pSheet->writeStr(0, 4, _T("本次") , pHeadFormat);
	// 内容
	int nRowIndex = 1;
	CString idx; 
	for (int i = 0; i < rules.GetCount(); i ++)
	{
		Crule rule = rules.GetAt(i);
		idx.Format(_T("%d"), (i+1));
		pSheet->writeStr(nRowIndex + i, 0, idx, pFormat);
		pSheet->writeStr(nRowIndex + i, 1, QualiClaimSheet.type, pFormat);
		pSheet->writeStr(nRowIndex + i, 2, rule.rule_name, pFormat);
		pSheet->writeStr(nRowIndex + i, 3, rule.rule_decribe, pFormat);
		if (rule.rule_use == _T("1"))
			pSheet->writeStr(nRowIndex + i, 4, _T("本次采用"), pFormat);
		else
			pSheet->writeStr(nRowIndex + i, 4, _T("本次不采用"), pFormat);
	}

	// 自动调整列宽
	int TextWidth_Max , TextWidth_Head , TextWidth_Final , TextWidth;
	CString Text ;
	int i,j;
	for ( i = 0; i < 100; i++)
	{
		for ( j=0; j< 50; j++)
		{
			if (j == 0)
			{ Text = pSheet->readStr(j, i, &pHeadFormat);
			TextWidth = Text.GetLength() ;
			TextWidth_Head = TextWidth;
			TextWidth_Max = TextWidth;}
			else
			{
				Text = pSheet->readStr(j, i, &pFormat);
				TextWidth = Text.GetLength() ;
			}
			if (TextWidth_Max < TextWidth)
				TextWidth_Max = TextWidth;

		}
		if (TextWidth_Max > 20)
			TextWidth_Max = 20;
		pSheet->setCol(i, i, TextWidth_Max *2, pDesignWidth, false);
	}
}


void excelTest(CQualiClaimSheet& QualiClaimSheet)
{
	libxl::Book *pBook = NULL;
	pBook = xlCreateBook();

	if (NULL == pBook)
	{
		::MessageBox(::GetActiveWindow(), _T("创建Excel对象失败！"), _T("提示"), MB_ICONERROR);
		return;
	}

	pBook->setKey(_T("yue liu"),
		_T("windows-2f21200806cae10968b46668a7hfe4l0"));

	//准备Excel格式
	// 无网格-列宽格式
	libxl::Format* pDesignWidth = pBook->addFormat();
	pDesignWidth->setWrap(false);
	// 细宋体加粗居中 pHeadFormat
	libxl::Format* pHeadFormat = pBook->addFormat();
	pHeadFormat->setAlignH(libxl::ALIGNH_CENTER);
	pHeadFormat->setAlignV(libxl::ALIGNV_CENTER);
	pHeadFormat->setWrap(true);
	pHeadFormat->setBorder(libxl::BORDERSTYLE_THIN);
	libxl::Font *pFont = pBook->addFont();
	pFont->setBold(true);
	pHeadFormat->setFont(pFont);
	// 细宋体居中 pFormat
	libxl::Format* pFormat = pBook->addFormat();
	pFormat->setAlignH(libxl::ALIGNH_CENTER);
	pFormat->setAlignV(libxl::ALIGNV_CENTER);
	pFormat->setWrap(true);
	pFormat->setBorder(libxl::BORDERSTYLE_THIN);
	// 细宋体靠上靠左 pLeftTopFormat
	libxl::Format* pLeftTopFormat = pBook->addFormat();
	pLeftTopFormat->setAlignH(libxl::ALIGNH_LEFT);
	pLeftTopFormat->setAlignV(libxl::ALIGNV_TOP);
	pLeftTopFormat->setWrap(true);
	pLeftTopFormat->setBorder(libxl::BORDERSTYLE_THIN);

	// 新建sheet
	libxl::Sheet *pSheet = NULL;

	ExportSheet1(pBook, pSheet, pHeadFormat, pFont, pFormat, pLeftTopFormat, pDesignWidth,QualiClaimSheet.AchiClaim.Arows, QualiClaimSheet.AchiClaim);
	ExportSheet2(pBook, pSheet, pHeadFormat, pFont, pFormat, pLeftTopFormat,pDesignWidth, QualiClaimSheet.ReptClaim.Rrows);
	ExportSheet3(pBook, pSheet, pHeadFormat, pFont, pFormat, pDesignWidth, QualiClaimSheet.cargos);
	ExportSheet4(pBook, pSheet, pHeadFormat, pFont, pFormat, pDesignWidth, QualiClaimSheet.infos, QualiClaimSheet.tnums,QualiClaimSheet);
	ExportSheet5(pBook, pSheet, pHeadFormat, pFont, pFormat, pDesignWidth, QualiClaimSheet.ReptClaim.result, QualiClaimSheet);
	ExportSheet6(pBook, pSheet, pHeadFormat, pFont, pFormat, pDesignWidth, QualiClaimSheet.rules, QualiClaimSheet);

	//保存文件
	if(!pBook->save(_T("D:\\sdkTest.xls")))
	{
		pBook->release();
		::MessageBox(::GetActiveWindow(), _T("Excel文件保存失败！"), _T("提示"), MB_ICONERROR);
		return;
	}

	pBook->release();
}