# Windows 10 21H2 Build 19044
# Visual Studio 2019
# Powershell $PSVersionTable 5.1

# HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualStudio.DTE\CLSID
# 2E1517DA-87BF-4443-984A-D2BF18F5A908

# Sources
# https://stackoverflow.com/questions/4724381/get-the-reference-of-the-dte2-object-in-visual-c-sharp-2010/4724924#4724924
# https://docs.microsoft.com/en-us/previous-versions/ms228755(v=vs.140)?redirectedfrom=MSDN
# https://blog.adamfurmanek.pl/2016/03/19/executing-c-code-using-powershell-script/

# Pass in solution name to return running VS DTE (from one of several running)
# Your calling process & Visual Studio(s) must run at same priviledge

Param( 
[Parameter(Mandatory=$true)]
[Alias("SLN")]
[ValidateNotNullOrEmpty()]
[string]$solutionName = "VSRunObjTest.sln"
)

$randForAssembly = Get-Random
$typeName = "JeffLomax$randForAssembly.Powershell"

$cs = @"
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace $typeName
{
    public class DTEUtil
    {
        [DllImport("ole32")]
        private static extern int CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid lpclsid);

        [DllImport("ole32.dll")]
        private static extern void CreateBindCtx(int reserved, out System.Runtime.InteropServices.ComTypes.IBindCtx ppbc);

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out System.Runtime.InteropServices.ComTypes.IRunningObjectTable prot);

        [DllImport("oleaut32")]
        private static extern int GetActiveObject([MarshalAs(UnmanagedType.LPStruct)] Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        public object GetActiveObject
        (
            string solutionFilename,
            string progId,
            bool throwOnError = false
        )
        {
            Guid clsid;
            object dteObject = null;

            if( progId == null )
            {
                throw new ArgumentNullException("progId");
            }

            int hr = CLSIDFromProgIDEx(progId, out clsid);
            if( hr < 0 )
            {
                if( throwOnError )
	            {
                    Marshal.ThrowExceptionForHR(hr);
	            }

                return null;
            }

            // Get RunningOjbectTable
            IRunningObjectTable rot;
            GetRunningObjectTable(0, out rot);
            if( rot == null )
            {
                throw new Exception("null IRunningObjectTable");
            }

            // Call EnumRunning
            IEnumMoniker monikerEnumerator;
            rot.EnumRunning(out monikerEnumerator);
            if( monikerEnumerator == null )
            {
                throw new Exception("null IEnumMoniker");
            }

            monikerEnumerator.Reset();
            IntPtr pNumFetched = new IntPtr();
            IMoniker[] monikers = new IMoniker[1];
            IMoniker moniker = null;
	        string displayName;

            while( dteObject==null && monikerEnumerator.Next(1, monikers, pNumFetched) == 0 )
	        {
                IBindCtx bindCtx;
                CreateBindCtx(0, out bindCtx);
		        if (bindCtx == null)
                {
                    continue;
                }

                moniker = monikers[0];
		        moniker.GetDisplayName(bindCtx, null, out displayName);

                object comObject = null;
                rot.GetObject(moniker, out comObject);
                if( comObject != null )
                {
                    // For Item Moniker "!VisualStudio.DTE.10.0:1234,", use ComObject
                    // For Solution Path Moniker, use ComObject.DTE
                    if( displayName.EndsWith( solutionFilename, StringComparison.CurrentCultureIgnoreCase ) )
                    {
                        dteObject = comObject;
                    }
                }

                Marshal.ReleaseComObject(bindCtx);
            }

            return dteObject;
        }
    }
}
"@

Add-Type -TypeDefinition $cs 

# VS2019 "VisualStudio.DTE.16.0"
$CurrentVisualStudioCLSID = "visualstudio.dte"

$dtecall = New-Object -TypeName "$typeName.DTEUtil"

$solutionRunningObject = $dtecall.GetActiveObject( $solutionName, $CurrentVisualStudioCLSID )

$dte = $solutionRunningObject.DTE

Write-Host "Active Document '$($dte.ActiveDocument.Name)'"
