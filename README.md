# GetMSLicenseList  获取系统有效证书

两种方式：  
通过调用SPPC实现  

 ```C
int GetLicenseList(System::Action<String^>^ GetResult) 
{

	String^ Name;
	String^ IID;
	String^ PartialProductKey;
	String^ Channel;
	String^ LicenseStatus;
	String^ StatusList;
	int ReArmSKUCount;
	List<String^>^ list = gcnew List<String^>();
	HSLC hSlc;
	if (SLOpen(&hSlc) == ERROR_SUCCESS)
	{

		SLDATATYPE peDataType;
		UINT pnReturnIds = 0;
		System::IntPtr ppbReturnIds;
		if (SLGetSLIDList(hSlc, SL_ID_PRODUCT_SKU, NULL, SL_ID_PRODUCT_SKU, &pnReturnIds, (SLID**)&ppbReturnIds) == 0)
		{
			for (int i = 0; i < pnReturnIds; i++)
			{
				SLID* pSkuId = (SLID*)malloc(sizeof(SLID));
				memcpy(pSkuId, (VOID*)(ppbReturnIds.ToInt64() + i * 0x10), 16);
				IntPtr pIID;
				if (SLGenerateOfflineInstallationId(hSlc, pSkuId, (PWSTR*)&pIID) == 0)
				{
					String^ IID = Marshal::PtrToStringUni(pIID);
					UINT pnProductKeyIds = 0;
					SLID* ppProductKeyIds;
					if (SLGetInstalledProductKeyIds(hSlc, pSkuId, &pnProductKeyIds, &ppProductKeyIds) == 0)
					{
						SLID* pkeyID = ppProductKeyIds;
						UINT pcbValue = 0;
						System::IntPtr ppbValue;
						if (SLGetPKeyInformation(hSlc, pkeyID, L"PartialProductKey", &peDataType, &pcbValue, (PBYTE*)&ppbValue) == 0)
						{
							PartialProductKey = Marshal::PtrToStringUni(ppbValue);
						}
						if (SLGetPKeyInformation(hSlc, pkeyID, L"Channel", &peDataType, &pcbValue, (PBYTE*)&ppbValue) == 0)
						{
							Channel = Marshal::PtrToStringUni(ppbValue)->Replace("Volume:", "")->Replace(":", "-");
						}
						if (SLGetProductSkuInformation(hSlc, pSkuId, L"Name", &peDataType, &pcbValue, (PBYTE*)&ppbValue) == 0)
						{
							Name = Marshal::PtrToStringUni(ppbValue)->ToString()->Split(',')[1]->Replace("edition", "")->Trim();
						}
						UINT pnStatusCount = 0;
						SLID* ppReturnIds;
						if (SLGetSLIDList(hSlc, SL_ID_PRODUCT_SKU, pSkuId, SL_ID_APPLICATION, &pnStatusCount, &ppReturnIds) == 0)
						{
							SLID* pAppId = ppReturnIds;
							if (IsEqualGUID(*pAppId, SLID_WINDOWS))
							{
								if (SLGetApplicationInformation(hSlc, pAppId, L"RemainingRearmCount", &peDataType, &pcbValue, (PBYTE*)&ppbValue) == 0)
								{
									ReArmSKUCount = (int)Marshal::ReadIntPtr(ppbValue);
								}
							}
							else
							{
								if (SLGetProductSkuInformation(hSlc, pSkuId, L"RemainingRearmCount", &peDataType, &pcbValue, (PBYTE*)&ppbValue) == 0)
								{
									ReArmSKUCount = (int)Marshal::ReadIntPtr(ppbValue);
								}
							}
							SL_LICENSING_STATUS* sLicensingStatus;
							if (SLGetLicensingStatusInformation(hSlc, pAppId, pSkuId, NULL, &pcbValue, &sLicensingStatus) == 0)
							{
								switch (sLicensingStatus->eStatus) {
								case SL_LICENSING_STATUS_IN_GRACE_PERIOD:
									LicenseStatus = GetOSLCID() == 1 ? "初始宽限期" : "OOBGrace";									
									break;
								case SL_LICENSING_STATUS_LAST:
									LicenseStatus = GetOSLCID() == 1 ? "超出宽限期" : "OOTGrace";									
									break;
								case SL_LICENSING_STATUS_LICENSED:									
									LicenseStatus = GetOSLCID() == 1 ? "已授权" : "Licensed";										
									break;
								case SL_LICENSING_STATUS_NOTIFICATION:
									LicenseStatus = GetOSLCID() == 1 ? "通知状态" : "ExtendedGrace";									
									break;
								case SL_LICENSING_STATUS_UNLICENSED:
									LicenseStatus = GetOSLCID() == 1 ? "未授权" : "Unlicensed";									
									break;
								default:
									LicenseStatus = GetOSLCID() == 1 ? "未知" : "UnKnown";									
									break;
								}
							}
						}
					}
					list->Add("Name:" + Name + "," + "Channel:" + Channel + "," + "Key:" + PartialProductKey + "," + "Activation:" + LicenseStatus + "," + "InstalltionId:" + IID);
					String^ s = "Name:" + Name + "," + "Channel:" + Channel + "," + "Key:" + PartialProductKey + "," + "Activation:" + LicenseStatus + "," + "InstalltionId:" + IID;					
					 myCallback(ss);		
				}
			}
		}
	}
	SLClose(hSlc);
}
```
