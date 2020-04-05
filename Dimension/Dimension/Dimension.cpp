// Dimension.cpp: 定义 DLL 的初始化例程。
//

#include "stdafx.h"
#include "Dimension.h"
#include "excel9.h"
#include <ProToolkit.h>
#include <ProUtil.h>
#include <ProArray.h>
#include <ProDimension.h>
#include <ProWindows.h>
#include <ProSolid.h>
#include <ProParameter.h>
#include <ProParamval.h>
#include <ProMessage.h>
#include <ProMenu.h>
#include <ProMenuBar.h>
#include <ProMdl.h>
#include <ProModelitem.h>
#include <ProDrawing.h>
#include <ProFeature.h>
#include <ProFeatType.h>
#include <ProDrawing.h>
#include <ProNotify.h>
#include <ProRefInfo.h>
#include <ProRelSet.h>
#include <ProSelection.h>
#include <ProUICmd.h>
#include <ProNotify.h>
#include <ProPopupmenu.h>
#include <ProSelbuffer.h>
#include <ProNote.h>
#include <ProAnnotation.h>
#include <ProAnnotationFeat.h>
#include <ProAnnotationElem.h>
#include <ProSurface.h>
#include <ProGtol.h>
#include <ProUIMessage.h>
#include <ProView.h>
#include <ProSurface.h>
#include <ProEdge.h>

/*
typedef enum ProErrors
{
	PRO_TK_NO_ERROR = 0,
	PRO_TK_GENERAL_ERROR = -1,
	PRO_TK_BAD_INPUTS = -2,
	PRO_TK_USER_ABORT = -3,
	PRO_TK_E_NOT_FOUND = -4,
	PRO_TK_E_FOUND = -5,
	PRO_TK_LINE_TOO_LONG = -6,
	PRO_TK_CONTINUE = -7,
	PRO_TK_BAD_CONTEXT = -8,
	PRO_TK_NOT_IMPLEMENTED = -9,
	PRO_TK_OUT_OF_MEMORY = -10,
	PRO_TK_COMM_ERROR = -11, // communication error通讯错误（连接错误）
	PRO_TK_NO_CHANGE = -12,
	PRO_TK_SUPP_PARENTS = -13,
	PRO_TK_PICK_ABOVE = -14,
	PRO_TK_INVALID_DIR = -15,
	PRO_TK_INVALID_FILE = -16,
	PRO_TK_CANT_WRITE = -17,
	PRO_TK_INVALID_TYPE = -18,
	PRO_TK_INVALID_PTR = -19,
	PRO_TK_UNAV_SEC = -20,
	PRO_TK_INVALID_MATRIX = -21,
	PRO_TK_INVALID_NAME = -22,
	PRO_TK_NOT_EXIST = -23,
	PRO_TK_CANT_OPEN = -24,
	PRO_TK_ABORT = -25,
	PRO_TK_NOT_VALID = -26,
	PRO_TK_INVALID_ITEM = -27,
	PRO_TK_MSG_NOT_FOUND = -28,
	PRO_TK_MSG_NO_TRANS = -29,
	PRO_TK_MSG_FMT_ERROR = -30,
	PRO_TK_MSG_USER_QUIT = -31,
	PRO_TK_MSG_TOO_LONG = -32,
	PRO_TK_CANT_ACCESS = -33,
	PRO_TK_OBSOLETE_FUNC = -34,
	PRO_TK_NO_COORD_SYSTEM = -35,
	PRO_TK_E_AMBIGUOUS = -36,
	PRO_TK_E_DEADLOCK = -37,
	PRO_TK_E_BUSY = -38,
	PRO_TK_E_IN_USE = -39,
	PRO_TK_NO_LICENSE = -40,
	PRO_TK_BSPL_UNSUITABLE_DEGREE = -41,
	PRO_TK_BSPL_NON_STD_END_KNOTS = -42,
	PRO_TK_BSPL_MULTI_INNER_KNOTS = -43,
	PRO_TK_BAD_SRF_CRV = -44,
	PRO_TK_EMPTY = -45,
	PRO_TK_BAD_DIM_ATTACH = -46,
	PRO_TK_NOT_DISPLAYED = -47,
	PRO_TK_CANT_MODIFY = -48,
	PRO_TK_CHECKOUT_CONFLICT = -49,
	PRO_TK_CRE_VIEW_BAD_SHEET = -50,
	PRO_TK_CRE_VIEW_BAD_MODEL = -51,
	PRO_TK_CRE_VIEW_BAD_PARENT = -52,
	PRO_TK_CRE_VIEW_BAD_TYPE = -53,
	PRO_TK_CRE_VIEW_BAD_EXPLODE = -54,
	PRO_TK_UNATTACHED_FEATS = -55,
	PRO_TK_REGEN_AGAIN = -56,
	PRO_TK_DWGCREATE_ERRORS = -57,
	PRO_TK_UNSUPPORTED = -58,
	PRO_TK_NO_PERMISSION = -59,
	PRO_TK_AUTHENTICATION_FAILURE = -60,
	PRO_TK_OUTDATED = -61,
	PRO_TK_INCOMPLETE = -62,
	PRO_TK_CHECK_OMITTED = -63,
	PRO_TK_MAX_LIMIT_REACHED = -64,
	PRO_TK_OUT_OF_RANGE = -65,
	PRO_TK_CHECK_LAST_ERROR = -66,

	
	//NOTE: the errors below are reserved for the Creo Toolkit API. Applications
	//should never return these errors.
	
	PRO_TK_APP_CREO_BARRED = -88,
	PRO_TK_APP_TOO_OLD = -89,
	PRO_TK_APP_BAD_DATAPATH = -90,
	PRO_TK_APP_BAD_ENCODING = -91,
	PRO_TK_APP_NO_LICENSE = -92,
	PRO_TK_APP_XS_CALLBACKS = -93,
	PRO_TK_APP_STARTUP_FAIL = -94,
	PRO_TK_APP_INIT_FAIL = -95,
	PRO_TK_APP_VERSION_MISMATCH = -96,
	PRO_TK_APP_COMM_FAILURE = -97,
	PRO_TK_APP_NEW_VERSION = -98,
	PRO_TK_APP_UNLOCK = -99,
	PRO_TK_APP_JLINK_NOT_ALLOWED = -100

}  ProError, ProErr;    // most commonly used Creo Parametric TOOLKIT error statuses
*/

ProError status;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

//
//TODO:  如果此 DLL 相对于 MFC DLL 是动态链接的，
//		则从此 DLL 导出的任何调入
//		MFC 的函数必须将 AFX_MANAGE_STATE 宏添加到
//		该函数的最前面。
//
//		例如: 
//
//		extern "C" BOOL PASCAL EXPORT ExportedFunction()
//		{
//			AFX_MANAGE_STATE(AfxGetStaticModuleState());
//			// 此处为普通函数体
//		}
//
//		此宏先于任何 MFC 调用
//		出现在每个函数中十分重要。  这意味着
//		它必须作为以下项中的第一个语句:
//		出现，甚至先于所有对象变量声明，
//		这是因为它们的构造函数可能生成 MFC
//		DLL 调用。
//
//		有关其他详细信息，
//		请参阅 MFC 技术说明 33 和 58。
//

// CDimensionApp

BEGIN_MESSAGE_MAP(CDimensionApp, CWinApp)
END_MESSAGE_MAP()


// CDimensionApp 构造

CDimensionApp::CDimensionApp()
{
	// TODO:  在此处添加构造代码，
	// 将所有重要的初始化放置在 InitInstance 中
}


// 唯一的 CDimensionApp 对象

CDimensionApp theApp;


// CDimensionApp 初始化

BOOL CDimensionApp::InitInstance()
{
	CWinApp::InitInstance();

	return TRUE;
}

//设置菜单在不同模式下的状态
static uiCmdAccessState AccessAvailable(uiCmdAccessMode access_mode)
{
	return (ACCESS_AVAILABLE);
}

/************************************************************************/
/* 获得当前模型          
 * typedef void* ProMdl;
*/
/************************************************************************/
ProMdl GetCurrentMdl()
{
	ProMdl     mdl;
	ProError   status;
	/*
   @UNDO: SAFE
   Purpose:  Initializes the <i>p_handle</i> with the current Creo Parametric
			 object.

   Input Arguments:
	  None

   Output Arguments:
	  p_handle     - The model handle

   Return Values:
	  PRO_TK_NO_ERROR    - The function successfully initialized the handle.
	  PRO_TK_BAD_CONTEXT - The current Creo Parametric object is not set.

	*/
	status = ProMdlCurrentGet(&mdl);
	if (status == PRO_TK_NO_ERROR)
		return mdl;
	return NULL;
}

/*
typedef enum{

		ACCESS_REMOVE		//移除菜单项
		ACCESS_INVISIBLE	//菜单项不可见
		ACCESS_UNAVAILABLE	//菜单项可见，变灰不可选
		ACCESS_DISALLOW		//菜单项不可选
		ASSESS_AVAILABLE	//菜单项可选
	}uiCmdAccessState
*/
static uiCmdAccessState UsrAccessDefault(uiCmdAccessMode access_mode)
{
	//如果当前存在模型且为零件模型则有效
	ProMdlType   p_type;
	if (GetCurrentMdl() != NULL)
	{
		/*
		 @UNDO: SAFE
		 Purpose: Returns the type of model (such as
					PRO_PART or PRO_ASSEMBLY).
	
		 Input Arguments:
			  model - A model pointer whose type needs to returned.

		 Output Arguments:
			  p_type - The type of model. If the function fails, this is
					   set to PRO_TYPE_UNUSED.

		 Return Values:
			  PRO_TK_NO_ERROR   - The function successfully retrieved the type
			  PRO_TK_BAD_INPUTS - The input argument is invalid.
		*/
		ProMdlTypeGet(GetCurrentMdl(), &p_type);
		if (p_type == PRO_MDL_PART)
		{
			return ACCESS_AVAILABLE;
		}
	}
	return ACCESS_UNAVAILABLE;

}

#define MFG_TEMPLATE_NAME L"MFG_AE_TEMPLATE_NAME"

//typedef wchar_t	ProLine[PRO_LINE_SIZE];
//typedef double  ProVector[3];
//typedef wchar_t	ProName[PRO_NAME_SIZE];
//typedef enum  ProBooleans
//{
//	PRO_B_FALSE = 0,
//	PRO_B_TRUE = 1
//} ProBoolean, ProBool;
typedef struct
{
	ProLine template_type;		//实际为wchar_t template_type[81]
	ProVector orientation;		//实际为double orientation[3]
	ProName orientation_label;	//实际为wchar_t orientation_label[32]
	ProBoolean is_completed;  /* runtime flag - indicates we processed this AE
							  into the BOM already */
} PTMfgTemplateOrientationData;

typedef struct
{
	/*
	typedef struct pro_model_item
		{
		  ProType  type;
		  int      id;
		  ProMdl owner;
		} ProModelitem, ProGeomitem, ProExtobj, ProFeature, ProProcstep,
		  ProSimprep, ProExpldstate, ProLayer, ProDimension, ProDtlnote,
		  ProDtlsyminst, ProGtol, ProCompdisp, ProDwgtable, ProNote,
		  ProAnnotationElem, ProAnnotation, ProAnnotationPlane, 
		  ProSymbol, ProSurfFinish, ProMechItem, ProMaterialItem, ProCombstate,
		  ProLayerstate, ProApprnstate;
	*/
	ProAnnotationElem ae;

	/*
	typedef enum 
		{
		    PRO_ANNOT_TYPE_NONE          = 0,
		    PRO_ANNOT_TYPE_NOTE          = PRO_NOTE,
		    PRO_ANNOT_TYPE_GTOL          = PRO_GTOL,
		    PRO_ANNOT_TYPE_SRFFIN        = PRO_SURF_FIN,
		    PRO_ANNOT_TYPE_SYMBOL        = PRO_SYMBOL_INSTANCE,
		    PRO_ANNOT_TYPE_DRVDIM        = PRO_DIMENSION,
		    PRO_ANNOT_TYPE_REFDIM        = PRO_REF_DIMENSION,
		    PRO_ANNOT_TYPE_SET_DATUM_TAG = PRO_SET_DATUM_TAG,
		    PRO_ANNOT_TYPE_CUSTOM        = PRO_CUSTOM_ANNOTATION,
		    PRO_ANNOT_TYPE_DRIVINGDIM    = PRO_ANNOT_ELEM_DRIVING_DIM
		} ProAnnotationType;
	*/
	ProAnnotationType type;

	int feat_id;
	void* app_data;
} PTAEInfo;

typedef ProError(*PTAEAppDataCollectionFunction) (ProAnnotationElem* ae, void** app_data);

typedef struct
{
	PTAEAppDataCollectionFunction app_collection_function;
	PTAEInfo* collected_aes;
} PTAECollectionData;

ProError PTMfgTemplateAEInfoCollect(ProAnnotationElem* ae, void** app_data)
{
	PTMfgTemplateOrientationData* data;
	/*
	typedef struct proparameter
		{
		  ProType       type;
		  ProName       id;  
		  ProParamowner owner;
		} ProParameter;
	*/
	ProParameter param;
	/*
	typedef struct  Pro_Param_Value  {
		 ProParamvalueType  	type;
		 ProParamvalueValue    value;
		}  ProParamvalue;
	*/
	ProParamvalue pvalue;

	data = (PTMfgTemplateOrientationData*)calloc(1, sizeof(PTMfgTemplateOrientationData));

	/*---------------------------------------------------------------------*\
	Extract the value of the template name parameter
	extern ProError ProParameterInit (	ProModelitem *owner,
										ProName       name,
										ProParameter *param);
	\*---------------------------------------------------------------------*/
	/*
		Purpose:	Initializes a <i>ProParameter</i> data structure.

		Input Arguments:
				owner		- The solid to which the <i>ProParameter</i> belongs
				name		- The name of the <i>ProParameter</i>

		Output Arguments:
				param		- The initialized <i>ProParameter</i> handle

		Return Values:
			PRO_TK_NO_ERROR		- The function successfully initialized the handle.
			PRO_TK_BAD_INPUTS	- One or more of the input arguments are invalid.
			PRO_TK_BAD_CONTEXT	- The owner is nonexistent.
			PRO_TK_E_NOT_FOUND	- The parameter was not found within the owner.
	*/
	status = ProParameterInit(ae, MFG_TEMPLATE_NAME, &param);

	status = ProParameterValueGet(&param, &pvalue);

	//extern  ProError  ProWstringCopy(wchar_t* source, wchar_t* target, int num_chars);
	/*
	   Purpose:
			 Copies a wide string into another buffer.
	   Input Arguments:
			 source - The source wide string.
			 target - The target wide string. It is the caller's responsibility to allocate enough memory for the copy operation.
			 num_chars -  The number of wide strings to copy. If PRO_VALUE_UNUSED, the entire string will be copied.
	   Output Arguments:
			none
	   Return Values:
		   PRO_TK_NO_ERROR - The information was returned successfully.
		   PRO_TK_BAD_INPUTS - One or more arguments was invalid.
	*/
	ProWstringCopy(pvalue.value.s_val, data->template_type, PRO_VALUE_UNUSED);

	/*---------------------------------------------------------------------*\
	Not implemented; for future use
	\*---------------------------------------------------------------------*/
	data->orientation[0] = 0.0;
	data->orientation[1] = 0.0;
	data->orientation[2] = 0.0;

	data->orientation_label[0] = (wchar_t)0;

	data->is_completed = PRO_B_FALSE;

	*app_data = data;

	return PRO_TK_NO_ERROR;
}

/*====================================================================*\
FUNCTION :   PTAEIsMfgTemplate()
PURPOSE  :   Identifies if the AE is a manufacturing template AE
\*====================================================================*/
ProError PTAEIsMfgTemplate(ProAnnotationElem* ae, ProAppData data)
{
	/*
	ProAnnotationType type;
	ProParameter param;

	status = ProAnnotationelemTypeGet (ae, &type);

	if (type != PRO_ANNOT_TYPE_CUSTOM)
	return PRO_TK_CONTINUE;

	status = ProParameterInit (ae, MFG_TEMPLATE_NAME, &param);

	if (status != PRO_TK_NO_ERROR)
	return PRO_TK_CONTINUE;

	*/
	return PRO_TK_NO_ERROR;
}

/*====================================================================*\
FUNCTION :   PTAEInfoCollect()
PURPOSE  :   Collect information on annotation elements
\*====================================================================*/
ProError PTAEInfoCollect(ProAnnotationElem* ae, ProError status, ProAppData data)
{
	PTAECollectionData* ae_data = (PTAECollectionData*)data;
	PTAEInfo info;
	ProFeature feat;

	/*---------------------------------------------------------------------*\
	Collect basic AE information: owner, id, feature that owns it
	\*---------------------------------------------------------------------*/
	info.ae.owner = ae->owner;
	info.ae.id = ae->id;
	info.ae.type = ae->type;

	status = ProAnnotationelemTypeGet(ae, &info.type);

	status = ProAnnotationelemFeatureGet(ae, &feat);

	info.feat_id = feat.id;

	/*---------------------------------------------------------------------*\
	Use the stored callback function to collect specific needed data
	\*---------------------------------------------------------------------*/
	if (ae_data->app_collection_function != NULL)
		(*ae_data->app_collection_function) (ae, (void**)&info.app_data);

	status = ProArrayObjectAdd((ProArray*)&ae_data->collected_aes, -1, 1, &info);

	return PRO_TK_NO_ERROR;
}

void outputinfo()
{
	//TCHAR szFilter[] = _T("Excel文件(*.xls)|*.xls|");
	TCHAR szFilter[] = _T("Excel文件|*.csv|Excel文件|*.xls|");
	//TCHAR szFilter[] = _T("*.xls");
	// 构造保存文件对话框   
	/*
	explicit CFileDialog(	BOOL bOpenFileDialog,			// TRUE for FileOpen, FALSE for FileSaveAs
							LPCTSTR lpszDefExt = NULL,
							LPCTSTR lpszFileName = NULL,
							DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
							LPCTSTR lpszFilter = NULL,
							CWnd* pParentWnd = NULL,
							DWORD dwSize = 0,
							BOOL bVistaStyle = TRUE);
	*/
	CFileDialog fileDlg(FALSE, _T("csv"), _T("Dimension information"), OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, szFilter);
	//CFileDialog fileDlg(FALSE, _T("csv"), _T("三维尺寸标注信息"), OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT);
	
	CString strFilePath;

	// 显示保存文件对话框   
	if (IDOK == fileDlg.DoModal())
	{
		strFilePath = fileDlg.GetPathName();

		CStringArray varName, varValue, DimType, DimUpperLimit, DimLowerLimit, Dimannoid, Dimrefid;
		ProMdl mdl;
		PTAECollectionData data;

		status = ProMdlCurrentGet(&mdl);

		data.app_collection_function = PTMfgTemplateAEInfoCollect;

		status = ProArrayAlloc(0, sizeof(PTAEInfo), 5, (ProArray*)&data.collected_aes);

		status = ProSolidAnnotationelemsVisit((ProSolid)mdl, PTAEInfoCollect, PTAEIsMfgTemplate, (ProAppData)&data);

		int num = 0;
		status = ProArraySizeGet((ProArray)data.collected_aes, &num);

		ProName symbol;
		double value;
		CString cvalue;
		ProDimensiontype type;
		CString indexNo;
		double upper_limit, lower_limit;
		CString cupper_limit, clower_limit;
		ProAnnotation dim_anno;
		CString coverall_value;
		CString cannoelemid;
		CString cv, st;

		for (int i = 0; i < num; i++)
		{
			if (data.collected_aes[i].type == PRO_ANNOT_TYPE_DRVDIM)
			{
				/*
				extern ProError ProAnnotationelemAnnotationSet (ProSelection element, ProAnnotation *annotation);

					 Purpose:  Gets the annotation contained in an annotation element.

					 Licensing Requirement: TOOLKIT for 3D Drawings

					 Input Arguments:
						element       - The annotation element.
					 Output Arguments:
						annotation    - The annotation.
					 Return Values:
						PRO_TK_NO_ERROR    - The function succeeded.
						PRO_TK_BAD_INPUTS  - One or more inputs was invalid.
						PRO_TK_E_NOT_FOUND - The annotation element does not contain an annotation.
				*/
				ProAnnotationelemAnnotationGet(&data.collected_aes[i].ae, &dim_anno);

				//尺寸符号
				ProDimensionSymbolGet(&dim_anno, symbol);
				varName.Add(symbol);

				//尺寸值
				ProDimensionValueGet(&dim_anno, &value);
				cvalue.Format(_T("%g"), value);
				varValue.Add(cvalue);

				//尺寸类型
				ProDimensionTypeGet(&dim_anno, &type);
				switch (type)
				{
				case PRODIMTYPE_LINEAR:
					DimType.Add(_T("DIMTYPE_LINEAR"));
					break;
				case PRODIMTYPE_RADIUS:
					DimType.Add(_T("DIMTYPE_RADIUS"));
					break;
				case PRODIMTYPE_DIAMETER:
					DimType.Add(_T("DIMTYPE_DIAMETER"));
					break;
				case PRODIMTYPE_ANGLE:
					DimType.Add(_T("DIMTYPE_ANGLE"));
					break;
				default:
					DimType.Add(_T("DIMTYPE_UNKNOWN"));
				}

				//上偏差、下偏差
				ProDimensionToleranceGet(&dim_anno, &upper_limit, &lower_limit);
				cupper_limit.Format(_T("%g"), upper_limit);
				DimUpperLimit.Add(cupper_limit);
				clower_limit.Format(_T("%g"), lower_limit);
				DimLowerLimit.Add(clower_limit);

				ProAnnotationReference *references = NULL;
				ProSelection ref_sel;
				ProModelitem ref_modelitem;
				CString ref_featname;

				int num_refs = 0;
				CUIntArray totalid;
				CStringArray DimFeatName;

				/*
				extern ProError ProAnnotationelemReferencesCollect( ProAnnotationElem        *element,
																	ProAnnotationRefFilter    ref_type,
																	ProAnnotationRefFromType  source,
																	ProAnnotationReference  **references);
				  Purpose:  Gets the references contained in an annotation element.

				  Licensing Requirement: TOOLKIT for 3D Drawings
				  Input Arguments:
					element       - The annotation element.
					ref_type      - The type of references to collect (weak, strong, or all).
					source        - The source of the references (from the annotation, from
									the user, or all).
				  Output Arguments:
					references    - The annotation references.  Free this using
									ProAnnotationreferencearrayFree().
				*/
				ProAnnotationelemReferencesCollect(&data.collected_aes[i].ae, PRO_ANNOTATION_REF_ALL, PRO_ANNOT_REF_FROM_ALL, &references);


				/*
				LIB_COREUTILS_API  ProError ProArraySizeGet (	ProArray array,
																int*     p_size);
				   Purpose:   Returns the size of the specified array.

				   Input Arguments:
					  array     - The array whose size is required
				   Output Arguments:
					  p_size    - The size of the array
				*/
				ProArraySizeGet((ProArray)references, &num_refs);

				for (int i = 0; i < num_refs; i++)
				{
					/*
					extern ProError ProReferenceToSelection (	ProReference  reference,
																ProSelection* selection);
					   Purpose: Gets and allocates a ProSelection containing a representation for this reference.
					   
					   Input Arguments:
							 reference          - The reference handle.
					   Output Arguments:
							 selection          - The resulting ProSelection handle.
												Note that this does not contain reference specific
												information like local copy flags, and status.
												This selection is separately allocated and
												should be freed using ProSelectionFree().
					*/
					ProReferenceToSelection(references[i].object.reference, &ref_sel);


					/*
					extern ProError ProSelectionModelitemGet(	ProSelection selection,
																ProModelitem *p_mdl_item );
					   Purpose: Gets the model item from a selection object.
					   Input Arguments:
						  selection  - The selection object.
					   Output Arguments:
						  p_mdl_item - The model item.
					*/
					ProSelectionModelitemGet(ref_sel, &ref_modelitem);

					totalid.Add(ref_modelitem.id);
					if (ref_modelitem.type == PRO_SURFACE)
					{
						ProSurface surface;
						ProGeomitemToSurface(&ref_modelitem, &surface);
						ProGeomitemdata* sur_data;
						ProSurfaceDataGet(surface, &sur_data);
						int sur_type = sur_data->data.p_surface_data->type;

						//平面
						if (sur_type == PRO_SRF_PLANE)
						{
							CString point1;
							point1.Format(_T("Length%g,Width%g"), sur_data->data.p_surface_data->uv_max[0], sur_data->data.p_surface_data->uv_max[1]);

							CString point2;
							point2.Format(_T("CentralPoint(%g,%g,%g)"), (sur_data->data.p_surface_data->xyz_max[0] + sur_data->data.p_surface_data->xyz_min[0]) / 2, (sur_data->data.p_surface_data->xyz_max[1] + sur_data->data.p_surface_data->xyz_min[1]) / 2, (sur_data->data.p_surface_data->xyz_max[2] + sur_data->data.p_surface_data->xyz_min[2]) / 2);

							CString point3;
							point3.Format(_T("InitialPoint(%g,%g,%g)"), sur_data->data.p_surface_data->srf_shape.plane.origin[0], sur_data->data.p_surface_data->srf_shape.plane.origin[1], sur_data->data.p_surface_data->srf_shape.plane.origin[2]);

							CString point4;
							point4.Format(_T("X[%g,%g],Y[%g,%g],Z[%g,%g]"), sur_data->data.p_surface_data->xyz_min[0], sur_data->data.p_surface_data->xyz_max[0], sur_data->data.p_surface_data->xyz_min[1], sur_data->data.p_surface_data->xyz_max[1], sur_data->data.p_surface_data->xyz_min[2], sur_data->data.p_surface_data->xyz_max[2]);

							CString c_sur_type;
							c_sur_type = _T("SRF_PLANE{") + point1 + _T(",") + point2 + _T(",") + point3 + _T(",") + point4 + _T("}");
							DimFeatName.Add(c_sur_type);
						}
						
						//圆柱面
						else if (sur_type == PRO_SRF_CYL)
						{
							CString point1;
							point1.Format(_T("Radius%g,Height%g"), sur_data->data.p_surface_data->srf_shape.cylinder.radius, sur_data->data.p_surface_data->uv_max[1]);

							CString point2;
							point2.Format(_T("CentralPoint(%g,%g,%g)"), (sur_data->data.p_surface_data->xyz_max[0] + sur_data->data.p_surface_data->xyz_min[0]) / 2, (sur_data->data.p_surface_data->xyz_max[1] + sur_data->data.p_surface_data->xyz_min[1]) / 2, (sur_data->data.p_surface_data->xyz_max[2] + sur_data->data.p_surface_data->xyz_min[2]) / 2);

							CString point3;
							point3.Format(_T("InitialPoint(%g,%g,%g)"), sur_data->data.p_surface_data->srf_shape.cylinder.origin[0], sur_data->data.p_surface_data->srf_shape.cylinder.origin[1], sur_data->data.p_surface_data->srf_shape.cylinder.origin[2]);

							CString point4;
							point4.Format(_T("X[%g,%g],Y[%g,%g],Z[%g,%g]"), sur_data->data.p_surface_data->xyz_min[0], sur_data->data.p_surface_data->xyz_max[0], sur_data->data.p_surface_data->xyz_min[1], sur_data->data.p_surface_data->xyz_max[1], sur_data->data.p_surface_data->xyz_min[2], sur_data->data.p_surface_data->xyz_max[2]);

							CString c_sur_type;
							c_sur_type = _T("SRF_CYL{") + point1 + _T(",") + point2 + _T(",") + point3 + _T(",") + point4 + _T("}");
							DimFeatName.Add(c_sur_type);
						}
						else
						{
							CString c_sur_type;
							c_sur_type = _T("SRF_OTHERS");
							DimFeatName.Add(c_sur_type);
						}
						ProGeomitemdataFree(&sur_data);
					}
					//边
					else if (ref_modelitem.type == PRO_EDGE)
					{
						ProEdge edge;
						ProGeomitemToEdge(&ref_modelitem, &edge);
						ProGeomitemdata* edge_data;
						ProEdgeDataGet(edge, &edge_data);

						double end10 = edge_data->data.p_curve_data->line.end1[0];
						double end11 = edge_data->data.p_curve_data->line.end1[1];
						double end12 = edge_data->data.p_curve_data->line.end1[2];

						double end20 = edge_data->data.p_curve_data->line.end2[0];
						double end21 = edge_data->data.p_curve_data->line.end2[1];
						double end22 = edge_data->data.p_curve_data->line.end2[2];

						if (end10<0.01 && end10>-0.01)
							end10 = 0;
						if (end11<0.01 && end11>-0.01)
							end11 = 0;
						if (end12<0.01 && end12>-0.01)
							end12 = 0;

						if (end20<0.01 && end20>-0.01)
							end20 = 0;
						if (end21<0.01 && end21>-0.01)
							end21 = 0;
						if (end22<0.01 && end22>-0.01)
							end22 = 0;

						CString point1;
						point1.Format(_T("EndPoint1(%g,%g,%g)"), end10, end11, end12);

						CString point2;
						point2.Format(_T("EndPoint2(%g,%g,%g)"), end20, end21, end22);

						CString c_sur_type;
						c_sur_type = _T("EDGE{") + point1 + _T(",") + point2 + _T("}");
						DimFeatName.Add(c_sur_type);

						ProGeomitemdataFree(&edge_data);
					}
					else
					{
						CString c_sur_type;
						c_sur_type = _T("SRF_OTHERS");
						DimFeatName.Add(c_sur_type);
					}
				}

				CString zong;

				if (num_refs == 1)
				{
					zong = DimFeatName[0];
				}
				else if (num_refs == 2)
				{
					zong = DimFeatName[0] + _T(",") + DimFeatName[1];
				}
				else if (num_refs == 3)
				{
					zong = DimFeatName[0] + _T(",") + DimFeatName[1] + _T(",") + DimFeatName[2];
				}
				else if (num_refs == 4)
				{
					zong = DimFeatName[0] + _T(",") + DimFeatName[1] + _T(",") + DimFeatName[2] + _T(",") + DimFeatName[3];
				}
				else if (num_refs == 5)
				{
					zong = DimFeatName[0] + _T(",") + DimFeatName[1] + _T(",") + DimFeatName[2] + _T(",") + DimFeatName[3] + _T(",") + DimFeatName[4];
				}
				else if (num_refs == 0)
				{
					zong = _T("None_Refs");
				}
				else
				{
					zong.Format(_T("%dRefs"), num_refs);
				}

				//class CStringArray   Dimrefid
				Dimrefid.Add(zong);


				/*
				extern ProError ProAnnotationreferencearrayFree (ProAnnotationReference* reference_array);
				  Purpose:  Frees all memory owned by the annotation reference array.
				  Input Arguments:
							reference_array     - The reference array.
				  Output Arguments:
							none
				*/
				ProAnnotationreferencearrayFree(references);


				/*
				extern ProError ProSelectionFree( ProSelection *p_selection );
				   Purpose: Frees a preallocated selection object.
				   Input Arguments:
							p_selection - The address of the selection object
				   Output Arguments:
							None
				*/
				ProSelectionFree(&ref_sel);
			}
		}


		_Application app;		//Excel应用程序接口
		Workbooks books;		//工作簿集合
		_Workbook book;		    //工作簿
		Worksheets sheets;		//工作表集合
		_Worksheet sheet;		//工作表
		Range range;			//Excel中针对单元格的操作都应先获取其对应的Range对象

		CString TempPath;
		//TempPath = _T("C:\\Dimension\\text\\template.xls");
		TempPath = _T("C:\\Dimension\\text\\template.csv");
		
		LPDISPATCH lpDisp;  //接口指针
		COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

		if (!app.CreateDispatch(_T("Excel.Application")))
		{
			AfxMessageBox(_T("无法创建Excel应用！"));
			return;
		}

		books.AttachDispatch(app.GetWorkbooks());
		lpDisp = books.Open(TempPath, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional);
		book.AttachDispatch(lpDisp);
		sheets.AttachDispatch(book.GetWorksheets());
		sheet.AttachDispatch(sheets.GetItem(COleVariant(long(1))));
		sheet.Activate();
		lpDisp = book.GetActiveSheet();
		sheet.AttachDispatch(lpDisp);
		range.AttachDispatch(sheet.GetCells(), TRUE);//加载所有单元格
		CString xuhao;

		for (int i = 0; i < varName.GetCount(); i++)
		{
			xuhao.Format(_T("%d"), i + 1);

			range.SetItem(COleVariant((long)(i + 2)), COleVariant((long)1), COleVariant(xuhao));

			range.SetItem(COleVariant((long)(i + 2)), COleVariant((long)2), COleVariant(varName[i]));

			range.SetItem(COleVariant((long)(i + 2)), COleVariant((long)3), COleVariant(DimType[i]));

			range.SetItem(COleVariant((long)(i + 2)), COleVariant((long)4), COleVariant(varValue[i]));

			range.SetItem(COleVariant((long)(i + 2)), COleVariant((long)5), COleVariant(DimUpperLimit[i]));

			range.SetItem(COleVariant((long)(i + 2)), COleVariant((long)6), COleVariant(DimLowerLimit[i]));

			range.SetItem(COleVariant((long)(i + 2)), COleVariant((long)7), COleVariant(Dimrefid[i]));
		}

		book.SaveAs(COleVariant(strFilePath), covOptional, covOptional, covOptional, covOptional, covOptional, (long)0, covOptional, covOptional, covOptional, covOptional); //与的不同，是个参数的，直接在后面加了两个covOptional成功了
		range.ReleaseDispatch();
		sheet.ReleaseDispatch();
		sheets.ReleaseDispatch();
		app.ReleaseDispatch();
		COleVariant aver((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		book.Close(aver, COleVariant(strFilePath), aver);
		books.Close();
		app.Quit();
		AfxMessageBox(_T("操作成功！"));
	}

	int CurrentWindowToActiveID;
	ProWindowCurrentGet(&CurrentWindowToActiveID);
	ProWindowActivate(CurrentWindowToActiveID);
}

/*
	函数：user_initialize()
	功能：用户初始化函数
*/
extern "C" int user_initialize()
{
	ProError status;
	ProFileName MsgFile;
	uiCmdCmdId   PushButton_cmd_id1;
	ProStringToWstring(MsgFile, "Message.txt");//设置菜单信息文件名

	/*=========================================================*\
	添加菜单条
	\*=========================================================*/

	/*
	函数作用：在Creo软件界面中添加一个新菜单

	ProError ProMenubarMenuAdd (ProMenuItemName menu_name, 菜单项名
								ProMenuItemLabel untranslated_menu_label, 菜单标签名
								ProMenuItemName neighbor, 相邻菜单名
								ProBoolean add_after_neighbor, 如果位于相邻菜单的右侧，则为PRO_B_TRUE;否则为左侧
								ProFileName filename, 菜单信息文件名)

	1.菜单项名在菜单体系中不能有相同的名称
	2.菜单标签名必须与信息文件中该段的标识关键字相同
	3.相邻菜单名不能为NULL
	*/
	status = ProMenubarMenuAdd("MainMenu", "Dimension", "Utilities", PRO_B_TRUE, MsgFile);

	/*=========================================================*\
	菜单按钮设置
	\*=========================================================*/

	/*
	函数作用：设置菜单项的动作

	ProError ProCmdActionAdd (char* action_name, 使用的动作命令名
							  uiCmdCmdActFn action_cb, 激活菜单时用的动作函数名
							  uiCmdPriority priority, 命令的优先级别
							  uiCmdAccessFn access_func, 确定菜单是否可选、不可选或隐藏的回调函数
							  ProBoolean allow_in_non_active_window, 布尔值，是否在非激活窗口显示该菜单项
							  ProBoolean allow_in_accessory_window, 布尔值，是否在附属窗口显示该菜单项
							  uiCmdCmdId* action_id, 动作函数的命令标志号)

	1.动作命令名必须是唯一的
	2.uiCmdPriority priority指命令的优先级别，可取预定义常数 0 2 3 5 6 7 999
	3.uiCmdAccessFn access_func包含的类型如下：
					typedef enum{

						ACCESS_REMOVE		//移除菜单项
						ACCESS_INVISIBLE	//菜单项不可见
						ACCESS_UNAVAILABLE	//菜单项可见，变灰不可选
						ACCESS_DISALLOW		//菜单项不可选
						ASSESS_AVAILABLE	//菜单项可选
					}uiCmdAccessState
	4.uiCmdCmdId* action_id是动作函数的命令标志号，在调用动作管理的ProMenubarmenuPushbuttonAdd函数时作为输入参数
	*/

	//设置菜单按钮1的动作函数
	ProCmdActionAdd("PushButtonAct1", (uiCmdCmdActFn)outputinfo, 
		uiCmdPrioDefault, (uiCmdAccessFn)UsrAccessDefault, PRO_B_TRUE, PRO_B_TRUE, &PushButton_cmd_id1);

	/*
	函数作用：在菜单中添加菜单按钮

	ProError ProMenubarmenuPushbuttonAdd (ProMenuItemName parent_menu, 父菜单名
										  ProMenuItemName push_button_name, 菜单名
										  ProMenuItemLabel push_button_label, 菜单标签名，该值必须与信息文件中同组的标识关键字相同
										  ProMenuLineHelp one_line_help, 菜单提示文本，该值必须与信息文件中同组的标识关键字相同
										  ProMenuItemName neighbor, 相邻菜单名
										  ProBooleam add_after_neighbor, 如果位于相邻菜单之后，则为PRO_B_TRUE;否则为之前
										  uiCmdCmdId action_id, 动作函数的命令标识号
										  ProFileName filename, 信息文件名)

	1.ProMenuItemName neighbor相邻菜单名若为NULL，则将该菜单添加至菜单的首项或最后一项
	*/

	//添加菜单1按钮
	ProMenubarmenuPushbuttonAdd("MainMenu", "Dimension extract", "output", 
		"Dimension extract", NULL, PRO_B_TRUE, PushButton_cmd_id1, MsgFile);

	return status;
}

/*
	编写信息文件：

	信息文件用来定义菜单项、菜单项提示等信息。有固定的格式，以4行为一组，含义如下：
	第1行：关键字，该关键字必须与使用该信息文件函数的相关字符串相同
	第2行：在菜单项或菜单项提示上显示的英语文本
	第3行：中文文本
	第4行：为空

	信息文件必须位于text文件夹下。

	本程序信息文件内容如下所示：

Dimension
Dimension extract
二次开发
#
output
output
尺寸提取
#
Dimension extract
Dimension extract
输出尺寸信息
#

*/

/*
	编写注册文件

	NAME			应用程序标识名
	EXEC_FILE		可执行程序名（包括路径）
	TEXT_DIR		text目录路径
	STARTUP			启动应用模式，dll
	ALLOW_STOP		若为TRUE，可在Creo工作时终止应用程序，否则不能终止
	DELAY_START		若为TRUE，Creo在启动时不调用应用程序，否则自动启动
	REVISION		Creo版本号
	END				结束标志

NAME test
EXEC_FILE C:\test\test\test.dll
TEXT_DIR C:\test\text
STARTUP    dll
REVISION   Creo4.0
END

*/

/*==========================================================*\
函数: user_terminate()
功能：用户结束中断函数
\*==========================================================*/
extern "C" void user_terminate()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
}