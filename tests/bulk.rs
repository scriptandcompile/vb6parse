use vb6parse::parsers::{VB6ClassFile, VB6FormFile, VB6ModuleFile, VB6Project};

#[test]
#[ignore]
fn bulk_load_all_projects() {
    let projects = [
        "./tests/data/VB6-add-GUI-objects-at-runtime/Generate objects at runtime V3/Project1.vbp",
        "./tests/data/VB6-add-GUI-objects-at-runtime/Generate objects at runtime V2/Project1.vbp",
        "./tests/data/VB6-add-GUI-objects-at-runtime/Generate objects at runtime V1/Project1.vbp",
        "./tests/data/OCX_Advanced_Control__VB6/ocxProject.vbp",
        "./tests/data/OCX_Advanced_Control__VB6/ButtonEx.vbp",
        "./tests/data/Prototype-software-for-Photon-pixel-coupling/Vesta.vbp",
        "./tests/data/Endless-runner-Game_VB6.0/Endless runner project files/game.vbp",
        "./tests/data/omelette-vb6/Samples/HelloWorld/HelloWorld.vbp",
        "./tests/data/omelette-vb6/Samples/Empty/Project1.vbp",
        "./tests/data/omelette-vb6/Omelette.vbp",
        "./tests/data/ChessBrainVB/ChessBrainVB_V4_03a/Source/ChessBrainVB_debug.vbp",
        "./tests/data/ChessBrainVB/ChessBrainVB_V4_03a/Source/ChessBrainVB_PCode.vbp",
        "./tests/data/ChessBrainVB/ChessBrainVB_V4_03a/Source/ChessBrainVB.vbp",
        "./tests/data/KORG_Read_pcg/PKorgTrReader.vbp",
        "./tests/data/Environment/testme.vbp",
        "./tests/data/Environment/mexe.vbp",
        "./tests/data/Environment/M2000.vbp",
        "./tests/data/Bitrate-calculator/Windows/Source-code/BitrateCalc.vbp",
        "./tests/data/SK-ADO_Dan_SQL_Demo__VB6/Project1.vbp",
        "./tests/data/OCX_Advanced_Grid__VB6/Project1.vbp",
        "./tests/data/Win_Dialogs/Classes/FolderBrowser/Projekt1.vbp",
        "./tests/data/Win_Dialogs/archiv/VBC_IFileDialog/Project1.vbp",
        "./tests/data/Win_Dialogs/archiv/fabel358/Projekt1.vbp",
        "./tests/data/Win_Dialogs/archiv/fabel358/orig/Projekt1.vbp",
        "./tests/data/Win_Dialogs/archiv/VBC_Tipp0759_TiKu/Projekt1.vbp",
        "./tests/data/Win_Dialogs/CDlgShowPrinter/Projekt1.vbp",
        "./tests/data/Win_Dialogs/FolderBrowser/vbarchiv/Projekt1.vbp",
        "./tests/data/Win_Dialogs/FolderBrowser/ActiveVB/Projekt1.vbp",
        "./tests/data/Win_Dialogs/FolderBrowser/SHGetPathFromIDList/Projekt1.vbp",
        "./tests/data/Win_Dialogs/FolderBrowser/ActiveVBW/Projekt1.vbp",
        "./tests/data/Win_Dialogs/FontDialog/codekabinett/Projekt1.vbp",
        "./tests/data/Win_Dialogs/PWinDialogs.vbp",
        "./tests/data/stdVBA-Inspiration/TrickAdvancedTools/TrickAdvancedTools.vbp",
        "./tests/data/stdVBA-Inspiration/CSuperCollection_Demo/PTestIEnumVARIANT.vbp",
        "./tests/data/stdVBA-Inspiration/DialogEx/Project1.vbp",
        "./tests/data/stdVBA-Inspiration/SaveLoadUDT/SaveLoadUDT.vbp",
        "./tests/data/stdVBA-Inspiration/WebEvents 2/WebEvents.vbp",
        "./tests/data/stdVBA-Inspiration/InjectingVB6 dll and manipulating UI/InjectAndManipulateVB-main/injector/Injector.vbp",
        "./tests/data/stdVBA-Inspiration/InjectingVB6 dll and manipulating UI/InjectAndManipulateVB-main/dll/GetVBGlobal.vbp",
        "./tests/data/stdVBA-Inspiration/InjectingVB6 dll and manipulating UI/InjectAndManipulateVB-main/dummy/Dummy.vbp",
        "./tests/data/stdVBA-Inspiration/VBProjectScanner/prjScanner.vbp",
        "./tests/data/stdVBA-Inspiration/AlphaImgControl/LaVolpeAlphaImg.vbp",
        "./tests/data/stdVBA-Inspiration/OlafSimpleCharting/SimpleCharting.vbp",
        "./tests/data/stdVBA-Inspiration/vbGraph/vbGraph.vbp",
        "./tests/data/stdVBA-Inspiration/vbGraph/Demo Project/GraphTest.vbp",
        "./tests/data/stdVBA-Inspiration/QuadTree/pQuadTree.vbp",
        "./tests/data/stdVBA-Inspiration/Gossamer 1-7/GossDemo1.vbp",
        "./tests/data/stdVBA-Inspiration/stringeval/StringEval.vbp",
        "./tests/data/stdVBA-Inspiration/prjUniDLLcalls/prjUniDLLcalls.vbp",
        "./tests/data/stdVBA-Inspiration/PNG to PictureBox/prj32bppSuite.vbp",
        "./tests/data/stdVBA-Inspiration/FauxInterface/Project1.vbp",
        "./tests/data/stdVBA-Inspiration/ListViewProgressBar/Proyecto1.vbp",
        "./tests/data/stdVBA-Inspiration/vbFriendlyInterfaces/TutorialApps/2 - LightWeight EarlyBound-Objects/Test.vbp",
        "./tests/data/stdVBA-Inspiration/vbFriendlyInterfaces/TutorialApps/8 - Simple SOAPDemo with vbIDispatch/SOAPVB6.vbp",
        "./tests/data/stdVBA-Inspiration/vbFriendlyInterfaces/TutorialApps/3 - LightWeight Object-Lists/Test.vbp",
        "./tests/data/stdVBA-Inspiration/vbFriendlyInterfaces/TutorialApps/7 - Dynamic usage of vbIDispatch/Test.vbp",
        "./tests/data/stdVBA-Inspiration/vbFriendlyInterfaces/TutorialApps/5 - MultiEnumerations per vbIEnumerable/Test.vbp",
        "./tests/data/stdVBA-Inspiration/vbFriendlyInterfaces/TutorialApps/9 - usage of vbIPictureDisp/Test.vbp",
        "./tests/data/stdVBA-Inspiration/vbFriendlyInterfaces/TutorialApps/1 - LightWeight LateBound-Objects/Test.vbp",
        "./tests/data/stdVBA-Inspiration/vbFriendlyInterfaces/TutorialApps/6 - Performance of vbIDispatch/Test.vbp",
        "./tests/data/stdVBA-Inspiration/vbFriendlyInterfaces/TutorialApps/0 - LightWeight COM without any Helpers/LightWeightCOM.vbp",
        "./tests/data/stdVBA-Inspiration/vbFriendlyInterfaces/TutorialApps/4 - Enumerables per vbIEnumVariant/Test.vbp",
        "./tests/data/stdVBA-Inspiration/vbFriendlyInterfaces/vbInterfaces-Dll/vbInterfaces.vbp",
        "./tests/data/stdVBA-Inspiration/MSO_UI_Editor/Proyecto1.vbp",
        "./tests/data/stdVBA-Inspiration/func_pointers_1.0.4/example/FuncPtrExample.vbp",
        "./tests/data/stdVBA-Inspiration/Library info 2/prjLibInfo.vbp",
        "./tests/data/stdVBA-Inspiration/vbActiveScript/vbActiveScript.vbp",
        "./tests/data/stdVBA-Inspiration/SpellCheck/Proyecto1.vbp",
        "./tests/data/stdVBA-Inspiration/MessageLog/prjMessageLog.vbp",
        "./tests/data/stdVBA-Inspiration/Better StringBuilder/Project1.vbp",
        "./tests/data/stdVBA-Inspiration/Simple-ActiveExe-InitThread/src/AxExeInitThread.vbp",
        "./tests/data/stdVBA-Inspiration/thunks2/Project1.vbp",
        "./tests/data/SK-Alarm_Clock__VB6/Project1.vbp",
        "./tests/data/CdiuBeatUpEditor/Project1.vbp",
        "./tests/data/Troyano-VB6-PoC/Client/Client.vbp",
        "./tests/data/Troyano-VB6-PoC/Server/Server.vbp",
        "./tests/data/Genomin/Project1.vbp",
        "./tests/data/ADM-TSC-Tools-ALM-QC/QC_RenameUsers/RenameUsers.vbp",
        "./tests/data/ProjectExaminer/ProjectExaminer.vbp",
        "./tests/data/NewTab/test in source code/TDIForms test/TDIFormsTest.vbp",
        "./tests/data/NewTab/test in source code/Test.vbp",
        "./tests/data/NewTab/control-source/NewTabCtl.vbp",
        "./tests/data/NewTab/test compiled ocx/TDIForms test/TDIFormsTest.vbp",
        "./tests/data/NewTab/test compiled ocx/Test.vbp",
        "./tests/data/Markov-Chains-VB6/Markov Chain V1.0/source/HMM.vbp",
        "./tests/data/Markov-Chains-VB6/Markov Chain V2.0/source/HMM.vbp",
        "./tests/data/Markov-Chains-VB6/Markov Chain V3.0/source/HMM.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/9.其他常用功能/9.1 全屏窗口切换/状态切换.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/5.综合测试/5.1 综合测试1/综合测试.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/2.文字/2.1 显示文字/文字.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/2.文字/2.2 旋转文字/文字旋转.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/3.输入设备/3.2 键盘检测/键盘.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/3.输入设备/3.1 鼠标检测/鼠标.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/3.输入设备/3.3 手柄检测/手柄.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/1.图像/1.2 贴图/贴图.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/1.图像/1.4 画线扩展/画线.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/1.图像/1.1 绘图/绘图.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/1.图像/1.3 光照特效/光照.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/4.音频/4.2 声音效果/动态变声.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/4.音频/4.3 空间音效/环绕音效.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/4.音频/4.1 播放声音/音频播放.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/开发示例/0.初始化/0.1 初始化引擎/初始化.vbp",
        "./tests/data/CoolWind2D-GameEngine-CHS/BAS模块/BAS模块.vbp",
        "./tests/data/ppdm/ppdm.vbp",
        "./tests/data/SK-SQL_Code_Generator_V2__VB6/SQLGenerator.vbp",
        "./tests/data/VB6-samples/tsEDIT-1.0/tsEDIT.vbp",
        "./tests/data/VB6-samples/tsIRCD-1.2/src/pre/MD5Hash/MD5Hash.vbp",
        "./tests/data/VB6-samples/tsIRCD-1.2/src/pre/SystrayIcon/Project1.vbp",
        "./tests/data/VB6-samples/tsIRCD-1.2/src/tsIRCd.vbp",
        "./tests/data/VB6-samples/tsNFO-1.0/TS-NFO.vbp",
        "./tests/data/VB6-samples/tsHTTPD/tsHTTPd.vbp",
        "./tests/data/VB6-samples/VB6-GRAPHICS/tsSINELINE-1.2/SineLine.vbp",
        "./tests/data/VB6-samples/VB6-GRAPHICS/ts3DMOTION-1.0/3DMotion.vbp",
        "./tests/data/VB6-samples/VB6-GRAPHICS/tsRAINDOTS-1.0/raindots.vbp",
        "./tests/data/VB6-samples/VB6-GRAPHICS/ts3DSTARS-1.0/3dstars10.vbp",
        "./tests/data/VB6-samples/VB6-GRAPHICS/tsTHEMATRIX-1.0/prjDesktop.vbp",
        "./tests/data/VB6-samples/tsMID-1.0/src/tsMID.vbp",
        "./tests/data/VB6-samples/tsCALC-1.0/tsCALC.vbp",
        "./tests/data/w.bloggar/Source/Localize/wb4Local.vbp",
        "./tests/data/w.bloggar/Source/wbloggar.vbp",
        "./tests/data/Mix-two-signals-by-using-Spectral-Forecast-in-VB6-app-v1.0/SF.vbp",
        "./tests/data/VB6-2D-Physic-Engine/prj2Dengine.vbp",
        "./tests/data/ucJLDatePicker/Demo.vbp",
        "./tests/data/VPN-Lifeguard/Windows/1.4.14/Source-code/VpnLifeguard.vbp",
        "./tests/data/Discrete-Probability-Detector-in-VB6/DPD.vbp",
        "./tests/data/Mix-two-signals-by-using-Spectral-Forecast-in-VB6-app-v2.0/SF.vbp",
        "./tests/data/project-duplication-detection-system/design and implementation of project duplication detection system/trapduplicate.vbp",
        "./tests/data/SteamyDock/SteamyDock.vbp",
        "./tests/data/vb6-code/Artificial-life/Artificial Life.vbp",
        "./tests/data/vb6-code/Blacklight-effect/Blacklight.vbp",
        "./tests/data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.vbp",
        "./tests/data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.vbp",
        "./tests/data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.vbp",
        "./tests/data/vb6-code/Color-shift-effect/ShiftColor.vbp",
        "./tests/data/vb6-code/Colorize-effect/Colorize.vbp",
        "./tests/data/vb6-code/Contrast-effect/Contrast.vbp",
        "./tests/data/vb6-code/Curves-effect/Curves.vbp",
        "./tests/data/vb6-code/Custom-image-filters/CustomFilters.vbp",
        "./tests/data/vb6-code/Diffuse-effect/Diffuse.vbp",
        "./tests/data/vb6-code/Edge-detection/EdgeDetection.vbp",
        "./tests/data/vb6-code/Emboss-engrave-effect/EmbossEngrave.vbp",
        "./tests/data/vb6-code/Fill-image-region/Fill_Region.vbp",
        "./tests/data/vb6-code/Fire-effect/FlameTest.vbp",
        "./tests/data/vb6-code/Game-physics-basic/Physics.vbp",
        "./tests/data/vb6-code/Gradient-2D/Gradient.vbp",
        "./tests/data/vb6-code/Grayscale-effect/Grayscale.vbp",
        "./tests/data/vb6-code/Hidden-Markov-model/HMM.vbp",
        "./tests/data/vb6-code/Histograms-advanced/Advanced Histograms.vbp",
        "./tests/data/vb6-code/Histograms-basic/Basic Histograms.vbp",
        "./tests/data/vb6-code/Levels-effect/Image Levels.vbp",
        "./tests/data/vb6-code/Mandelbrot/Mandelbrot.vbp",
        "./tests/data/vb6-code/Map-editor-2D/Map Editor.vbp",
        "./tests/data/vb6-code/Nature-effects/NatureFilters.vbp",
        "./tests/data/vb6-code/Randomize-effects/RandomizationFX.vbp",
        "./tests/data/vb6-code/Scanner-TWAIN/VB_Scanner_Support.vbp",
        "./tests/data/vb6-code/Screen-capture/ScreenCapture.vbp",
        "./tests/data/vb6-code/Sepia-effect/Sepia.vbp",
        "./tests/data/vb6-code/Threshold-effect/Threshold.vbp",
        "./tests/data/vb6-code/Transparency-2D/Transparency.vbp",
        "./tests/data/vb6/MakeZhC/MakeZhC.vbp",
        "./tests/data/vb6/lUseZip/TestZip.vbp",
        "./tests/data/vb6/lUseZip/lUseZip.vbp",
        "./tests/data/vb6/HtmlParser/htmlParser.vbp",
        "./tests/data/vb6/ArtCorridor/Source/ArtCorridor.vbp",
        "./tests/data/vb6/StopZZZ/StopZZZ.Core.vbp",
        "./tests/data/vb6/Fac/Fac.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TNotify.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TPaths.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/RegTlbOld.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/CollWiz.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/LotAbout.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/GlobWiz.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Components/Notify.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Components/SubTimer.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Components/VBCore.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Components/VisualCore.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TPalette.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Timeit.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TIcon.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Sieve/SieveBasCtlN.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Sieve/SieveBasExeN.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Sieve/SieveBasDllP.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Sieve/SieveBasDllN.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Sieve/SieveBasGlobalN.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Sieve/SieveBasExeP.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Sieve/SieveBasGlobalP.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Sieve/SieveBasCtlP.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TImage.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TWindow.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Hardcore.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TRes2.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TSort.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TSplit2.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/AddrOMatic.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Bugwiz.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TExecute.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TSplit.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TWhiz.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/AppPath.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TMessage.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TBezier.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Browse.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TShortcut.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TEdge.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TParse.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TCollect.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/SieveCli.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/BitBlast.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TCompletion.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TColorPick.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TRes.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/VB6ToVB5.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Controls/PictureGlass.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Controls/ColorPicker.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Controls/DropStack.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Controls/ListBoxPlus.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Controls/Editor.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TTimer.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TThread.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/RegTlb.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TSysMenu.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TFolder.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TDictionary.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Winwatch.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/AllAbout.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TShare.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/ErrMsg.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Meriwether.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TEnum.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/TString.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/Edwina.vbp",
        "./tests/data/vb6/[Include]/HardcodeVB/FunNGame.vbp",
        "./tests/data/vb6/[Include]/Controls/xrzControls.vbp",
        "./tests/data/vb6/xrzUnpackDir/xrzUnpackDir.vbp",
        "./tests/data/vb6/dllFolder_Browser/vbalFolderBrowse6.vbp",
        "./tests/data/vb6/cleaner/cleaner.vbp",
        "./tests/data/vb6/zpShower/Source/zpShower.vbp",
        "./tests/data/vb6/LiNNetwork/LiNNetWork.vbp",
        "./tests/data/vb6/TestVSS/Project1.vbp",
        "./tests/data/vb6/MyMusic/MyMusic.vbp",
        "./tests/data/vb6/lContextMenu/LContextMenu.vbp",
        "./tests/data/vb6/FolderPacker/FolderPacker.vbp",
        "./tests/data/vb6/CreateFolderIndex/IndexDir.vbp",
        "./tests/data/vb6/BlogText/BlogSave.vbp",
        "./tests/data/vb6/BlogText/BlogWrite.vbp",
        "./tests/data/vb6/netCanoe/netCanoe.vbp",
        "./tests/data/vb6/PartitionTextFile/PartitionTextFile.vbp",
        "./tests/data/vb6/pdgLibQuery/pdgLibQuery.vbp",
        "./tests/data/vb6/LoadWow/LoadWow.vbp",
        "./tests/data/vb6/toolKitIDLgen/toolkitIDLgen.vbp",
        "./tests/data/vb6/addin_openDir/vbadding_openDir.vbp",
        "./tests/data/vb6/[Samples]/vbhash/Demo1/Demo.vbp",
        "./tests/data/vb6/CBookUrl/mpBookUrl.vbp",
        "./tests/data/vb6/XUnpack/xUnpack.vbp",
        "./tests/data/vb6/FolderBrowser/FolderBrowser.vbp",
        "./tests/data/vb6/RssDll/Rss.vbp",
        "./tests/data/vb6/txtReader/txtReader.vbp",
        "./tests/data/vb6/WenXin/wenxin.vbp",
        "./tests/data/vb6/LiNSubTimer/SubTimer.vbp",
        "./tests/data/vb6/NextPage/NextPageApp.vbp",
        "./tests/data/vb6/NextPage/NextPageDll.vbp",
        "./tests/data/vb6/GlobalWizard/GlobWiz.vbp",
        "./tests/data/vb6/zhReader/Source/zhReader.vbp",
        "./tests/data/vb6/zhReader/Source/zhSReader.vbp",
        "./tests/data/vb6/zhReader/Source/zhPReader.vbp",
        "./tests/data/vb6/lBlueSky/lBlueSky.vbp",
        "./tests/data/vb6/LExplorer/LExplorer.vbp",
        "./tests/data/vb6/ProgramLoader/ProgramLoader.vbp",
        "./tests/data/vb6/oevbext/OEVBEXT.vbp",
        "./tests/data/vb6/ssMdbQuery/ssMDBQuery.vbp",
        "./tests/data/vb6/FormatBlog/FormatBlog.vbp",
        "./tests/data/vb6/Fileinstr/Fileinstr.vbp",
        "./tests/data/vb6/sslibExplorer/sslibExplorer.vbp",
        "./tests/data/vb6/LTest/LTest.vbp",
        "./tests/data/vb6/IeSaveText/AppIST.vbp",
        "./tests/data/vb6/IeSaveText/DllIST.vbp",
        "./tests/data/vb6/91reg/91reg.vbp",
        "./tests/data/vb6/ExecLine/prjExecLine.vbp",
        "./tests/data/vb6/LiNVBLib/TestLiNVBLiB.vbp",
        "./tests/data/vb6/LiNVBLib/LiNVBLib.vbp",
        "./tests/data/vb6/lin-zip/lin-zip.vbp",
        "./tests/data/vb6/Delpdg/Deletepdg.vbp",
        "./tests/data/vb6/DevFormat/DevStudio6Format.vbp",
        "./tests/data/vb6/iebho/IEBHO.vbp",
        "./tests/data/vb6/p/p.vbp",
        "./tests/data/vb6/zipLoader/zipLoader.vbp",
        "./tests/data/vb6/chmMake/chmmake.vbp",
        "./tests/data/vb6/FirefoxPortableLoader/FirefoxPortableLoader.vbp",
        "./tests/data/vb6/zipProtocolBeta/zipProtocol.vbp",
        "./tests/data/vb6/fgbookUrl/bookUrl.vbp",
        "./tests/data/vb6/fgbookUrl/fgBookUrl.vbp",
        "./tests/data/vb6/FolderInstr/FolderInstr.vbp",
        "./tests/data/vb6/Santa/Santa.vbp",
        "./tests/data/vb6/textPdgMerger/textPdgMerger.vbp",
        "./tests/data/vb6/TimeCMD/TimeCMD.vbp",
        "./tests/data/vb6/PackChm/PackChm.vbp",
        "./tests/data/vb6/BinTree/BINTREE.vbp",
        "./tests/data/vb6/LiNControls/LiNControls.vbp",
        "./tests/data/vb6/Packpdg/packpdg.vbp",
        "./tests/data/vb6/SSDownload/smdh.vbp",
        "./tests/data/vb6/ssLibQuery/ssLibQuery.vbp",
        "./tests/data/vb6/VBSourceWizard/GlobWiz.vbp",
        "./tests/data/vb6/SyncDirectory/SyncDirectory.vbp",
        "./tests/data/vb6/InvisibleRun/InvisibleRun.vbp",
        "./tests/data/vb6/practiceCOM/practiceCOM.vbp",
        "./tests/data/vb6/Rename pdg folders/Rename PDG folders.vbp",
        "./tests/data/vb6/pdgzf/pdgZF.vbp",
        "./tests/data/vb6/ssLibBase/ssLibBase.vbp",
        "./tests/data/vb6/QuickWork/quickwork.vbp",
        "./tests/data/vb6/FileStrReplace/FileStrReplace.vbp",
        "./tests/data/vb6/Netcat/Netcat.vbp",
        "./tests/data/vb6/bookUrl/pbookUrl.vbp",
        "./tests/data/vb6/bookUrl/bookUrl.vbp",
        "./tests/data/vb6/ClassTemplate/TemplateType.vbp",
        "./tests/data/vb6/ClassTemplate/ClassTemplate.vbp",
        "./tests/data/vb6/ClassTemplate/ClassBuilder/ClassBuilder.vbp",
        "./tests/data/vb6/rssbho/RSSBHO.vbp",
        "./tests/data/vb6/[Template]/Projects/VB Enterprise Edition Controls.vbp",
        "./tests/data/vb6/[Template]/Projects/ActiveX Dll.vbp",
        "./tests/data/vb6/[Template]/Projects/LEmptyProject.vbp",
        "./tests/data/vb6/[Template]/Projects/Addin.vbp",
        "./tests/data/vb6/[Template]/Projects/Data Project.vbp",
        "./tests/data/vb6/[Template]/Projects/IEBHO.vbp",
        "./tests/data/vb6/[Template]/Projects/Activex Document Exe.vbp",
        "./tests/data/vb6/[Template]/Projects/IIS Application.vbp",
        "./tests/data/vb6/[Template]/Projects/ActiveX Document Dll.vbp",
        "./tests/data/vb6/[Template]/Projects/LProject.vbp",
        "./tests/data/vb6/[Template]/Projects/DHTML Application.vbp",
        "./tests/data/vb6/GetssLib/TaskMan.vbp",
        "./tests/data/vb6/GetssLib/复件 GetSSLIB.vbp",
        "./tests/data/vb6/GetssLib/LibGetSSLib.vbp",
        "./tests/data/vb6/GetssLib/GetSSLIB.vbp",
        "./tests/data/vb6/GetssLib/GetSSLIBx.vbp",
        "./tests/data/vb6/GetssLib/SSLibTaskman/Context/GetSSLibContext.vbp",
        "./tests/data/vb6/GetssLib/SSLibTaskman/SSLibTaskman.vbp",
        "./tests/data/vb6/ArchReader/ARCore.vbp",
        "./tests/data/vb6/ArchReader/ArchReader.vbp",
        "./tests/data/vb6/IncludeAll/IncludeAll.vbp",
        "./tests/data/vb6/IncludeAll/ClassAndModule.vbp",
        "./tests/data/vb6/IncludeAll/MakeIncludeAll.vbp",
        "./tests/data/vb6/Lookout/LookOut.vbp",
        "./tests/data/vb6/zipProtocol/zipProtocol.vbp",
        "./tests/data/vb6/BookManager/BookManager.vbp",
        "./tests/data/PromKappa-1.0-makes-Objective-Digital-Stains/source/CG.vbp",
        "./tests/data/SK-Password-Application-ADD-ON__VB6/Password.vbp",
        "./tests/data/VbScalesReader/src/Project1.vbp",
        "./tests/data/opendialup/programs/discador/Proyecto1.vbp",
        "./tests/data/audiostation/Audiostation/Audiostation.vbp",
        "./tests/data/Binary-metamorphosis/tini/tini.vbp",
        "./tests/data/Binary-metamorphosis/Binary metamorphosis V2.0/src/Bin_To_VB.vbp",
        "./tests/data/Binary-metamorphosis/Binary metamorphosis V3.0/src/Bin_To_VB.vbp",
        "./tests/data/Binary-metamorphosis/Binary metamorphosis V1.0/src/Bin_To_VB.vbp",
        "./tests/data/unlightvbe_qs/UnlightVBE-QS/UL.vbp",
    ];

    println!("Loading projects...");

    for project_path in projects.iter() {
        println!("Loading project: {}", project_path);

        let project_contents = std::fs::read(project_path).unwrap();

        let project_file_name = std::path::Path::new(project_path)
            .file_name()
            .unwrap()
            .to_str()
            .unwrap();

        let project = VB6Project::parse(project_file_name, project_contents.as_slice());

        if project.is_err() {
            println!(
                "Failed to load project '{}'\r\n{}",
                project_path,
                project.err().unwrap()
            );
            continue;
        }

        let project = project.unwrap();

        //remove filename from path
        let project_directory = std::path::Path::new(project_path).parent().unwrap();

        for class_reference in project.classes {
            let class_path = project_directory.join(&class_reference.path.to_string());

            if std::fs::metadata(&class_path).is_err() {
                println!(
                    "{} | Class not found: {}",
                    project_path,
                    class_path.to_str().unwrap()
                );
                continue;
            }

            println!("Loading class: {}", class_path.to_str().unwrap());

            let file_name = class_path.file_name().unwrap().to_str().unwrap();
            let class_contents = std::fs::read(&class_path).unwrap();
            let class = VB6ClassFile::parse(file_name.to_owned(), &mut class_contents.as_slice());

            if class.is_err() {
                println!(
                    "{} | Class load error: {}",
                    project_path,
                    class.err().unwrap()
                );
                continue;
            }

            let _class = class.unwrap();
        }

        for module_reference in project.modules {
            let module_path = project_directory.join(&module_reference.path.to_string());

            if std::fs::metadata(&module_path).is_err() {
                println!(
                    "{} | Module not found: {}",
                    project_path,
                    module_path.to_str().unwrap()
                );
                continue;
            }

            println!("Loading module: {}", module_path.to_str().unwrap());

            let file_name = module_path.file_name().unwrap().to_str().unwrap();
            let module_contents = std::fs::read(&module_path).unwrap();
            let module = VB6ModuleFile::parse(file_name.to_owned(), &module_contents);

            if module.is_err() {
                println!(
                    "{} | Module load error: {}",
                    project_path,
                    module.err().unwrap()
                );
                continue;
            }

            let _module = module.unwrap();
        }

        for form_reference in project.forms {
            let form_path = project_directory.join(&form_reference.to_string());

            if std::fs::metadata(&form_path).is_err() {
                println!(
                    "{} | Form not found: {}",
                    project_path,
                    form_path.to_str().unwrap()
                );
                continue;
            }

            println!("Loading form: {}", form_path.to_str().unwrap());

            let file_name = form_path.file_name().unwrap().to_str().unwrap();
            let form_contents = std::fs::read(&form_path).unwrap();
            let form = VB6FormFile::parse(file_name.to_owned(), &mut form_contents.as_slice());

            if form.is_err() {
                println!(
                    "{} | Form load error: {}",
                    project_path,
                    form.err().unwrap()
                );
                continue;
            }

            let _form = form.unwrap();
        }

        println!("Project loaded: {}", project_path);
    }
}
