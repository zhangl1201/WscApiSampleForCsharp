using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace WscApiSampleForCSharp
{
    public class WscProductInfoAgent
    {
        private Guid _clsidWSCProductList = new Guid("17072F7B-9ABE-4A74-A261-1EB76B55107A");
        private List<WSC_SECURITY_PROVIDER> _providers = new List<WSC_SECURITY_PROVIDER>()
        {
            WSC_SECURITY_PROVIDER.WSC_SECURITY_PROVIDER_ANTISPYWARE,
            WSC_SECURITY_PROVIDER.WSC_SECURITY_PROVIDER_ANTIVIRUS,
            WSC_SECURITY_PROVIDER.WSC_SECURITY_PROVIDER_FIREWALL
        };

        public void CheckForSecurityCenterProducts()
        {
            foreach (var provider in this._providers)
            {
                IWSCProductList pWSCProductList;
                HRESULT hr = CheckForWscProductList(provider, out pWSCProductList);

                Console.WriteLine($"Security Provider type : {provider}");

                if (HRESULT.S_OK == hr)
                {
                    uint nProductCount = 0;
                    hr = pWSCProductList.get_Count(out nProductCount);
                    if (HRESULT.S_OK == hr)
                    {
                        for (uint i = 0; i < nProductCount; i++)
                        {
                            IWscProduct pWscProduct;
                            hr = pWSCProductList.get_Item(i, out pWscProduct);
                            if (HRESULT.S_OK == hr)
                            {
                                string sProductName = new string('\0', 260);
                                string sRemediationPath = new string('\0', 260);
                                string sProductStateTimestamp = new string('\0', 260);
                                WSC_SECURITY_PRODUCT_STATE nProductState = 0;
                                WSC_SECURITY_SIGNATURE_STATUS nProductStatus;

                                if (HRESULT.S_OK == pWscProduct.get_ProductName(out sProductName))
                                {
                                    Console.WriteLine($"Product Name : {sProductName}");
                                }
                                if (HRESULT.S_OK == pWscProduct.get_RemediationPath(out sRemediationPath))
                                {
                                    Console.WriteLine($"Remediation Path : {sRemediationPath}");
                                }
                                if (HRESULT.S_OK == pWscProduct.get_ProductState(out nProductState))
                                {
                                    Console.WriteLine($"Product state : {nProductState}");
                                }
                                if (provider != WSC_SECURITY_PROVIDER.WSC_SECURITY_PROVIDER_FIREWALL
                                    && HRESULT.S_OK == pWscProduct.get_SignatureStatus(out nProductStatus))
                                {
                                    Console.WriteLine($"Product Signature Status : {nProductStatus}");
                                }
                                if (provider == WSC_SECURITY_PROVIDER.WSC_SECURITY_PROVIDER_ANTIVIRUS 
                                    && HRESULT.S_OK == pWscProduct.get_ProductStateTimestamp(out sProductStateTimestamp))
                                {
                                    Console.WriteLine($"Product state time stamp : {sProductStateTimestamp}");
                                }
                                Console.WriteLine();
                                Marshal.ReleaseComObject(pWscProduct);
                            }
                        }
                    }
                    Marshal.ReleaseComObject(pWSCProductList);
                }
                Console.WriteLine();

            }
        }

        private HRESULT CheckForWscProductList(WSC_SECURITY_PROVIDER provider, out IWSCProductList pWSCProductList)
        {
            Type WSCProductListType = Type.GetTypeFromCLSID(this._clsidWSCProductList, true);
            object WSCProductList = Activator.CreateInstance(WSCProductListType);
            pWSCProductList = (IWSCProductList)WSCProductList;
            HRESULT hr = pWSCProductList.Initialize((uint)provider);

            return hr;
        }
    }
}
