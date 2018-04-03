#include "stdafx.h"

HRESULT LoadRuntime_4(ICorRuntimeHost*&);

int _tmain(int argc, TCHAR* argv[]) {
  errno_t err = 0;
  TCHAR file_path[MAX_PATH] = {0};
  bool has_error = false;
  if (argc != 2) {
    has_error = true;
  } else {
    err = _tcscat_s(file_path, argv[1]);
    if (err != 0) has_error = true;
  }
  if (has_error) {
    std::cout << std::endl;
    std::cout << "Usage: tinyshell.exe <path_to_ps1_file>" << std::endl;
    std::cout << "Developed by: @dev4coffee" << std::endl;
    std::cout << std::endl;
    return 1;
  }

  HRESULT hr = S_OK;
  ICorRuntimeHost* host_ptr = nullptr;
  IUnknownPtr app_domain_thunk = NULL;
  _AppDomainPtr app_domain = NULL;
  bstr_t assembly_name(
      _T("System.Management.Automation, Version=1.0.0.0, Culture=neutral, ")
      _T("PublicKeyToken=31bf3856ad364e35"));
  _AssemblyPtr assembly_ptr = nullptr;
  bstr_t powershell_name(_T("System.Management.Automation.PowerShell"));
  _TypePtr powershell_type = nullptr;
  bstr_t func_create(_T("Create"));
  bstr_t func_add_script(_T("AddScript"));
  bstr_t func_invoke(_T("Invoke"));
  _MemberInfoPtr mem_invoke_ptr = nullptr;
  _MethodInfoPtr invoke_ptr = nullptr;
  BSTR bstr_output;

  variant_t vt_empty;
  variant_t vt_powershell;
  variant_t vt_script_path(file_path);
  variant_t vt_outputs;
  variant_t vt_output;

  SAFEARRAY* args_null = SafeArrayCreateVector(VT_VARIANT, 0, 0);
  SAFEARRAY* args_script = SafeArrayCreateVector(VT_VARIANT, 0, 1);
  SAFEARRAY* arr_methods = nullptr;
  LONG index = 0;
  hr = SafeArrayPutElement(args_script, &index, &vt_script_path);
  if (FAILED(hr)) goto ReleaseRuntime;

  // First, we try to load .NET runtime 4.0
  // If failed, try to load .NET runtime 2.0
  hr = LoadRuntime_4(host_ptr);
  if (FAILED(hr))
    hr = CorBindToRuntimeEx(_T("v2.0.50727"), _T("wks"), 0,
                            CLSID_CorRuntimeHost, IID_PPV_ARGS(&host_ptr));
  if (FAILED(hr)) goto ReleaseRuntime;

  // Start .NET runtime
  hr = host_ptr->Start();
  if (FAILED(hr)) goto ReleaseRuntime;

  // Get default domain
  hr = host_ptr->GetDefaultDomain(&app_domain_thunk);
  if (FAILED(hr)) goto ReleaseRuntime;
  app_domain_thunk->QueryInterface(IID_PPV_ARGS(&app_domain));
  if (FAILED(hr)) goto ReleaseRuntime;

  // Load automation.dll and powershell type
  hr = app_domain->Load_2(assembly_name, &assembly_ptr);
  if (FAILED(hr)) goto ReleaseRuntime;
  assembly_ptr->GetType_2(powershell_name, &powershell_type);
  if (FAILED(hr)) goto ReleaseRuntime;

  // Create instance and invoke command
  hr = powershell_type->InvokeMember_3(func_create, (BindingFlags)280, nullptr,
                                       vt_empty, args_null, &vt_powershell);
  if (FAILED(hr)) goto ReleaseRuntime;
  hr = powershell_type->InvokeMember_3(func_add_script, (BindingFlags)276,
                                       nullptr, vt_powershell, args_script,
                                       nullptr);
  if (FAILED(hr)) goto ReleaseRuntime;
  hr = powershell_type->GetMember(func_invoke, MemberTypes_Method,
                                  (BindingFlags)20, &arr_methods);
  if (FAILED(hr)) goto ReleaseRuntime;
  index = 0;
  hr = SafeArrayGetElement(arr_methods, &index, &mem_invoke_ptr);
  if (FAILED(hr)) goto ReleaseRuntime;
  hr = mem_invoke_ptr->QueryInterface(IID_PPV_ARGS(&invoke_ptr));
  if (FAILED(hr)) goto ReleaseRuntime;
  hr = invoke_ptr->Invoke_3(vt_powershell, args_null, &vt_outputs);
  if (FAILED(hr)) goto ReleaseRuntime;

  // Get output and print to console
  hr = ((ICollectionPtr)vt_outputs.punkVal)->get_Count(&index);
  if (FAILED(hr)) goto ReleaseRuntime;
  for (long i = 0; i < index; i++) {
    hr = ((IListPtr)vt_outputs.punkVal)->get_Item(i, &vt_output);
    if (FAILED(hr)) continue;
    if (vt_output.vt == VT_BSTR) {
      std::wcout << vt_output.bstrVal << std::endl;
    } else {
      hr = ((_ObjectPtr)vt_output.punkVal)->get_ToString(&bstr_output);
      if (FAILED(hr)) continue;
      std::wcout << bstr_output << std::endl;
    }
  }

ReleaseRuntime:
  if (host_ptr) {
    host_ptr->Stop();
    host_ptr->Release();
  }
  if (FAILED(hr)) return ResultHandler(hr);

  return 0;
}

HRESULT LoadRuntime_4(ICorRuntimeHost*& host_ptr) {
  HRESULT hr = S_OK;

  ICLRMetaHost* meta_ptr = nullptr;
  ICLRRuntimeInfo* info_ptr = nullptr;
  BOOL is_loadable = FALSE;

  hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&meta_ptr));
  if (FAILED(hr)) goto ErrorLoadRuntime_4;

  hr = meta_ptr->GetRuntime(_T("v4.0.30319"), IID_PPV_ARGS(&info_ptr));
  if (FAILED(hr)) goto ErrorLoadRuntime_4;

  hr = info_ptr->IsLoadable(&is_loadable);
  if (FAILED(hr)) goto ErrorLoadRuntime_4;
  if (!is_loadable) {
    hr = -1;
    goto ErrorLoadRuntime_4;
  }

  hr = info_ptr->GetInterface(CLSID_CorRuntimeHost, IID_PPV_ARGS(&host_ptr));

ErrorLoadRuntime_4:
  if (meta_ptr) meta_ptr->Release();
  if (info_ptr) info_ptr->Release();

  return hr;
}