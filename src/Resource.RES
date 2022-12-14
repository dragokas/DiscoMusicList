        ??  ??                  tB  $   # 2 4   ??     0        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">

    <assemblyIdentity
        version="1.0.0.0"
        processorArchitecture="*"
        name="AlexDragokas.DiscoList"
        type="win32">
    </assemblyIdentity>

    <description>Tool for creating config file for Disco Mod plugin by MitchDizzle</description>

    <dependency>
        <dependentAssembly>
            <assemblyIdentity
                type="win32"
                name="Microsoft.Windows.Common-Controls"
                version="6.0.0.0"
                processorArchitecture="X86"
                publicKeyToken="6595b64144ccf1df"
                language="*"
             />
        </dependentAssembly>
    </dependency>

    <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
        <security>
            <requestedPrivileges>
                <requestedExecutionLevel
                    level="asInvoker"
                    uiAccess="false">
                </requestedExecutionLevel>
            </requestedPrivileges>
        </security>
    </trustInfo>

	<!-- For mscomctl.ocx Product version 6.01.9545 -->
	<file name="mscomctl.ocx">
	    <comClass description="Microsoft ListView Control 6.0 (SP6)" clsid="{BDD1F04B-858B-11D1-B16A-00C0F0283628}" threadingModel="Apartment" progid="MSComctlLib.ListViewCtrl" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	    <comClass description="Microsoft TreeView Control 6.0 (SP6)" clsid="{C74190B6-8589-11D1-B16A-00C0F0283628}" threadingModel="Apartment" progid="MSComctlLib.TreeCtrl" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	    <comClass description="Microsoft TabStrip Control 6.0 (SP6)" clsid="{1EFB6596-857C-11D1-B16A-00C0F0283628}" threadingModel="Apartment" progid="MSComctlLib.TabStrip" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	    <comClass description="Microsoft Toolbar Control 6.0 (SP6)" clsid="{66833FE6-8583-11D1-B16A-00C0F0283628}" threadingModel="Apartment" progid="MSComctlLib.Toolbar" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	    <comClass description="Microsoft ImageList Control 6.0 (SP6)" clsid="{2C247F23-8591-11D1-B16A-00C0F0283628}" threadingModel="Apartment" progid="MSComctlLib.ImageListCtrl" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	    <comClass description="Microsoft StatusBar Control 6.0 (SP6)" clsid="{8E3867A3-8586-11D1-B16A-00C0F0283628}" threadingModel="Apartment" progid="MSComctlLib.SBarCtrl" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	    <comClass description="Microsoft Slider Control 6.0 (SP6)" clsid="{F08DF954-8592-11D1-B16A-00C0F0283628}" threadingModel="Apartment" progid="MSComctlLib.Slider" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	    <comClass description="Microsoft ProgressBar Control 6.0 (SP6)" clsid="{35053A22-8589-11D1-B16A-00C0F0283628}" threadingModel="Apartment" progid="MSComctlLib.ProgCtrl" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	    <comClass description="Microsoft ImageComboBox Control 6.0 (SP6)" clsid="{DD9DA666-8594-11D1-B16A-00C0F0283628}" threadingModel="Apartment" progid="MSComctlLib.ImageComboCtl" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	    <comClass description="TreeView General Property Page Object" clsid="{C27CCE32-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="TabStrip General Property Page Object" clsid="{C27CCE33-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="Tab Property Page Object" clsid="{C27CCE34-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="Toolbar General Property Page Object" clsid="{C27CCE37-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="Button Property Page Object" clsid="{C27CCE38-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="ImageList General Property Page Object" clsid="{C27CCE35-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="Image Property Page Object" clsid="{C27CCE36-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="StatusBar General Property Page Object" clsid="{C27CCE39-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="Panel Property Page Object" clsid="{C27CCE3A-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="Slider General Property Page Object" clsid="{C27CCE3C-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="Slider Appearance Property Page Object" clsid="{C27CCE3D-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="Progress Bar General Property Page Object" clsid="{C27CCE3B-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="ListView General Property Page Object" clsid="{C27CCE3E-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="ListView Sort Property Page Object" clsid="{C27CCE3F-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="ListView Images Property Page Object" clsid="{C27CCE40-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="ListView Columns Property Page Object" clsid="{C27CCE41-8596-11D1-B16A-00C0F0283628}" />
	    <comClass description="ImageComboBox General Property Page Object" clsid="{C27CCE42-8596-11D1-B16A-00C0F0283628}" />
	    <typelib tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" version="2.0" helpdir="" />
	</file>
	
	<!-- For mscomctl.ocx Product version 6.01.9545 -->
	<comInterfaceExternalProxyStub name="DImageComboEvents" iid="{DD9DA665-8594-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020420-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IButton" iid="{66833FE9-8583-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IButtonMenu" iid="{66833FED-8583-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IButtonMenus" iid="{66833FEB-8583-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IButtons" iid="{66833FE7-8583-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IColumnHeader" iid="{BDD1F051-858B-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IColumnHeaders" iid="{BDD1F050-858B-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IComboItem" iid="{DD9DA660-8594-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IComboItems" iid="{DD9DA662-8594-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IControls" iid="{C8A3DC00-8593-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IImage" iid="{2C247F26-8591-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IImageCombo" iid="{DD9DA664-8594-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IImageList" iid="{2C247F21-8591-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IImages" iid="{2C247F24-8591-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IListItem" iid="{BDD1F04E-858B-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IListItems" iid="{BDD1F04C-858B-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IListSubItem" iid="{BDD1F055-858B-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IListSubItems" iid="{BDD1F053-858B-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IListView" iid="{BDD1F049-858B-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="ImageListEvents" iid="{2C247F22-8591-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020420-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="INode" iid="{C74190B8-8589-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="INodes" iid="{C74190B7-8589-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IPanel" iid="{8E3867AA-8586-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IPanels" iid="{8E3867A4-8586-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IProgressBar" iid="{35053A20-8589-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IProgressBarEvents" iid="{35053A21-8589-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020420-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="ISlider" iid="{F08DF952-8592-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="ISliderEvents" iid="{F08DF953-8592-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020420-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IStatusBar" iid="{8E3867A1-8586-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IStatusBarEvents" iid="{8E3867A2-8586-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020420-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="ITab" iid="{1EFB6599-857C-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="ITabs" iid="{1EFB6597-857C-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="ITabStrip" iid="{1EFB6594-857C-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="ITabStripEvents" iid="{1EFB6595-857C-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020420-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IToolbar" iid="{66833FE4-8583-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="IToolbarEvents" iid="{66833FE5-8583-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020420-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="ITreeView" iid="{C74190B4-8589-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="ITreeViewEvents" iid="{C74190B5-8589-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020420-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	<comInterfaceExternalProxyStub name="ListViewEvents" iid="{BDD1F04A-858B-11D1-B16A-00C0F0283628}" proxyStubClsid32="{00020420-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
	
    <!-- Used in most (if not all) of the above files. -->
    <comInterfaceExternalProxyStub name="IVBDataObject" iid="{2334D2B1-713E-11CF-8AE5-00AA00C00905}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />
    <comInterfaceExternalProxyStub name="IVBDataObjectFiles" iid="{2334D2B3-713E-11CF-8AE5-00AA00C00905}" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" />

    <compatibility xmlns="urn:schemas-microsoft-com:compatibility.v1">
        <application>
            <!-- Windows Vista -->
            <supportedOS Id="{e2011457-1546-43c5-a5fe-008deee3d3f0}"/>
            <!-- Windows 7 -->
            <supportedOS Id="{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"/>
            <!-- Windows 8 -->
            <supportedOS Id="{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"/>
            <!-- Windows 8.1 -->
            <supportedOS Id="{1f676c76-80e1-4239-95bb-83d0f6d0da78}"/>
            <!-- Windows 10 -->
            <supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/>
        </application>
    </compatibility>

    <application xmlns="urn:schemas-microsoft-com:asm.v3">
        <windowsSettings>
            <dpiAware xmlns="http://schemas.microsoft.com/SMI/2005/WindowsSettings">true</dpiAware>
        </windowsSettings>
    </application>

</assembly>   