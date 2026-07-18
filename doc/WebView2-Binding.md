# [VB6] WebView2 Binding (Edge/Chromium)

Source: <https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)>

Archived: 2026-07-18 (14 pages)

## #1 Schmidt — Nov 1st, 2020, 08:56 AM

**VB6 WebView2-Binding (Edge-Chromium)**

Have just finished a Binding to the new WebView2-BrowserControl (based on Edge-Chromium).

 I've included this Binding (all in a single Class, named cWebView2) in the new RC6-version of the RichClient-lib
 (please download this new version 6 from its usual place, at vbRichClient.com first).

 The new BaseDll-package of the RC6 now includes the official WebView2Loader.dll (version 1.0.674),
 which the cWebView2-class then works against.

 Please note, that the above Binding will currently require, that you install the larger:
 "Evergreeen WebView2-Runtime" (not included in the RC6-BasePackage).
 Here the official MS-DownloadLink for the evergreen-installer: <https://go.microsoft.com/fwlink/p/?LinkId=2124703> https://go.microsoft.com/fwlink/p/?LinkId=2124703

 So, after ensuring the mentioned two prerequisites:
 - the Dlls of the new RC6-package in a folder of your choice + a registered RC6.dll
 - and the successfull installation of the "evergreen-WebView2-runtime" via the MS-download-link above

 You should now be able to test this new Edge-Browser-Binding (even on Win7-OSes) via this little VB6-Demo:
 <https://www.vbforums.com/attachment.php?s=9b84d96bf559669a1dc45b4f80d160b8&attachmentid=179166&d=1604238779> WebView2Demo.zip

 Please let me know, when something is not working as expected -
 or when you want me to include a certain extra-functionality into the new cWebView2-class.

 I want to "finalize" the new RC6-functionality at the end of the year (then switching Binary-Compatibility on).

 Happy testing... *[img: Smilie]*

 Olaf

## #2 xiaoyao — Nov 2nd, 2020, 08:10 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

VERY GOOD,HOW TO get web page cookie? so i can use winhttp or xmlhttp to download URL
 OR WebView2Loader.dll WILL have a method for download url,post data?

## #3 xiaoyao — Nov 2nd, 2020, 08:15 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

CodeName=Mozilla
 MinorVersion=undefined
 Name=Netscape
 Version=5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36 Edg/86.0.622.58
 CookieEnabled=true

 how to test web page,its chrome or ie?
 <https://www.w3school.com.cn/tiy/t.asp?f=hdom_navigator> https://www.w3school.com.cn/tiy/t.asp?f=hdom_navigator

## #4 xiaoyao — Nov 2nd, 2020, 09:10 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

HOW TO RUN JS CALL BY WEB page like :var a=btn1_click(22,33)?
 i want to call vb sub from js,and back the value to js

Code:

```
    WV.AddScriptToExecuteOnDocumentCreated "function btn1_click(msg1,msg2){ vbH().RaiseMessageEvent('btn1_click',msg1+','+msg2) }"

WV.NavigateToString "<!DOCTYPE

html><html><head><title>AppTitle</title></head><body>" &

_
                          "<div>Hello World...</div>" & _
                          "<input id='txt1' value='foo'>" & _
                          "<button id='btn1' onclick='btn1_click(22,33)'

>Button1</button>" & _
                      "</body></html>"
Private Sub WV_JSMessage(ByVal sMsg As String, ByVal

sMsgContent As String, oJSONContent As cCollection)
sMsgContent=22,33
```

## #5 SearchingDataOnly — Nov 2nd, 2020, 10:03 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

(From dm)

 I tested WebView2Demo and it's great. I have some questions:

 (1) Could you add IndexByItem to RC6.cCollection?

 (2) Can WebView.CapturePreview capture the entire Web page, not just the visible part of the window?

 (3) Could you demonstrate the drag operations of WebView, for example: Drag a button or grid from the VB6-Form to WebView, and WebView will generate a Html-Button or a Html-Table in the Web page. Similar to the following: <https://www.layoutit.com/build> https://www.layoutit.com/build

 (4) Can V8 be integrated into RC6?

 (5) If V8 is too large, can RC6 integrate a small embedded JavaScript engine similar to QuickJS or MuJS?

 (6) When we use VB6 to develop commercial controls, .NET developers are not willing to use VB6 controls, they prefer to use .NET controls. When we use VB6 to develop commercial software, our commercial software can only run on the Windows platform, while our competitors' products can run on multiple platforms (Windows, Web, Android, iOS). In other words, we don't know what we can do with VB6 now? If RC6 could help developers in cross-platform, it would be a very wonderful thing.

## #6 Schmidt — Nov 2nd, 2020, 12:02 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

HOW TO get web page cookie?

The WebView2-interfaces do **not **include (other than the old, IE-based WebControl)
 a COM-Wrapper-ObjectModel for interaction with the DOM of a loaded document.

 So, all the interaction with a currently loaded DOM has to happen via JavaScript.
 That's the reason, why:
 - jsProp(...)
 - jsRun(...)
 - jsRunAsync(...)
 are included as Methods behind the WV-variable (of type cWebView2).

 And in case of Cookies, that means, you will have to address them via the appropriate js-expression like this:

Code:

```
'enhance the Forms WebView-EventHandler below about an additional line, to read a given documents cookies
'when pressing the < Navigate to google.com > Button, you should see the cookies which were set from the Google-Server.
Private Sub WV_DocumentComplete()
  Debug.Print "WV_DocumentComplete"
  Debug.Print WV.jsProp("document.cookie")
End Sub
```

As for your BrowserVersion-String - the one you've posted:
 ...Version=5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36 Edg/86.0.622.58
 does contain a "Chromium-Version-Part" and an "Edge-Version-Part" as well... (so it's definitely not IE-based)

 Olaf

## #7 Schmidt — Nov 2nd, 2020, 01:30 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

(1) Could you add IndexByItem to RC6.cCollection?

This is already implemented (via Optional ByRef Param: FoundItemIndex) in the cCollection.ItemExists(...) method.

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

(2) Can WebView.CapturePreview capture the entire Web page, not just the visible part of the window?

One could determine (via JavaScript) the "total Height in Pixels" a Document would fit into
 (whilst having a certain PixelWidth) - and then manually resize the WebView (via WV.Movexxx) temporarily,
 and then call the Capture-Method.

 But this feature-request came up also in the "Issue-tracker" IIRC, so I'd like to wait for the official release.

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

(3) Could you demonstrate the drag operations of WebView, for example: Drag a button or grid from the VB6-Form to WebView, and WebView will generate a Html-Button or a Html-Table in the Web page. Similar to the following: <https://www.layoutit.com/build> https://www.layoutit.com/build

As already statet, one has to incorporate ones own js-snippets into the WebView2-Control,
 to address such actions.
 As e.g. shown here: <https://www.w3schools.com/html/html5_draganddrop.asp> https://www.w3schools.com/html/html5_draganddrop.asp

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

(4) Can V8 be integrated into RC6?

With the WebView2-Binding, one already has access to the V8-js-engine (it is included via Chromium within this new Edge-based Control).

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

(5) If V8 is too large, can RC6 integrate a small embedded JavaScript engine similar to QuickJS or MuJS?

I'd think that JScript9 (via cActiveScript) is already a "small and relatively fast" js-Engine one can use -
 and as said, for V8-support, one can use the WebView2 (e.g. hosted-in/bound-to an invisible PicBox).

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

(6) When we use VB6 to develop commercial controls, .NET developers are not willing to use VB6 controls, they prefer to use .NET controls. When we use VB6 to develop commercial software, our commercial software can only run on the Windows platform, while our competitors' products can run on multiple platforms (Windows, Web, Android, iOS). In other words, we don't know what we can do with VB6 now? If RC6 could help developers in cross-platform, it would be a very wonderful thing.

That depends on, when a VB6-compatible compiler becomes available.
 My spare-time is not sufficient, to drive this side-project of mine along in a decent pace currently.

 Meanwhile others might come up with "something promising" in this regard -
 but the timeframe will in either case be measured in years, not months.

 Olaf

## #8 LeandroA — Nov 3rd, 2020, 04:34 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf, I couldn't make it run

 Err Number -2147221164 Clase no registrada

 Set WV = New_c.WebView2 'create the instance

 what am I doing wrong?
 i think i can't register rc6.dll, but if i can load it to project

## #9 Schmidt — Nov 3rd, 2020, 08:27 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **LeandroA** *[img: View Post]*

what am I doing wrong?
 i think i can't register rc6.dll, but if i can load it to project

I think, that's what you need to resolve first.
 Just loading the typelib of the RC6 (semiautomatically... via the VB6-IDE),
 will not be enough (on a developer-machine).

 You'll have to:
 - keep any of the Dlls of the RC6-BasePackage-Dlls "in one place" (a single Folder you copy them to, as e.g. C:\RC6\)
 - and then "deep-register" the RC6.dll there either via the included VBScript or via an **Admin**-Console calling regsvr32 behind the SysWow64 folder

 HTH

 Olaf

## #10 SearchingDataOnly — Nov 3rd, 2020, 12:08 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

(From dm)

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

This is already implemented (via Optional ByRef Param: FoundItemIndex) in the cCollection.ItemExists(...) method.

Yes, in many cases, cCollection.ItemExists and FoundItemIndex are more convenient. But IndexByItem is more concise and intuitive, it's very necessary to add IndexByItem to cCollection.

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

I'd think that JScript9 (via cActiveScript) is already a "small and relatively fast" js-Engine one can use -
 and as said, for V8-support, one can use the WebView2 (e.g. hosted-in/bound-to an invisible PicBox).

cActiveScript does not seem to support some of the latest JS syntax and the latest RE parameters.

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

That depends on, when a VB6-compatible compiler becomes available.
 My spare-time is not sufficient, to drive this side-project of mine along in a decent pace currently.

If a programming language wants to gain long-term vitality, it must have good commercial value, so as to form a stable software ecological chain around this programming language. If RC5/RC6 users can develop some valuable commercial software around RC6, then you will have more resources (funds and teams) to speed up the development of VB6-compatible compiler. I don't know if you have had such a plan.

 I can develop some good commercial software with VB6 and RC5/RC6, but the problem is that if a commercial software can only run in a Windows environment, its market demand and market share will be very low. Such softwares cannot form a solid software ecological chain, nor can they provide you with more resources.

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Meanwhile others might come up with "something promising" in this regard -
 but the timeframe will in either case be measured in years, not months.

A few years is too long.
 I'm thinking, if we develop a scripting language compatible with the syntax of the BASIC language, and then develop a scripting language runtime environment similar to NodeJS. Is it possible to realize a VB6-like programming language that can be cross-platform? Of course, this scripting language must have a visual IDE similar to VB6.

## #11 Arnoutdv — Nov 3rd, 2020, 02:43 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Check RadBasic, Mercury VB, B4X
 But this has nothing to do with this code submission

## #12 wqweto — Nov 4th, 2020, 12:18 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

A few years is too long..

Like everyhting in life a new VB compiler will take enough time, patience and disposable income. Judging by the expressed lack of time *and* obvious lack of patience you must be in posession of a lot of disposable income.

 Is there something that interested parties must know here?

 cheers,
 </wqw>

## #13 SearchingDataOnly — Nov 7th, 2020, 11:36 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

A small question about WebView2Demo:

 After starting the computer, when opening WebView2Demo.exe or WebView2Demo.vbp for the first time, the system prompts "couldn't initialize Webview-Bingding". After closing the window and re-opening WebView2Demo.exe or WebView2Demo.vbp, the system runs normally.

 In other words, every time the computer is restarted, the first initialization (WV.BindTo) always fails, and the second time is normal.

 In addition, how does IE-WebBrowser achieve a function similar to WebView2.CapturePreview?

 (OS: Win10)

## #14 Schmidt — Nov 8th, 2020, 04:36 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

...every time the computer is restarted, the first initialization (WV.BindTo) always fails, and the second time is normal.

The BindTo (as well as the two Navigate-Methods) have a second (optional)TimeOutSeconds-Parameter -
 by default sitting at 5 seconds.

 I guess your system is (shortly after booting) still quite busy with delayed loading of other base-stuff like services etc. -
 in addition to loading the quite large Chromium-Runtime, ... so the default of 5sec may be a bit too short -
 (I can imagine, that older magnetic disks might play a role, because I see no larger delay here on my test-OSes on SSDs).

 You could:
 - either try to increase that TimeOut-Param (to maybe 10sec or so)
 - or set it to Zero, to tell the methods to wait for their matching EventHandlers
 .. (WV_InitComplete or WV_NavigationCompleted respectively, to proceed with your own, now async Inits from within those Handlers)

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

In addition, how does IE-WebBrowser achieve a function similar to WebView2.CapturePreview?

You still planning on using the IE further???

 Well, there's an IHTMLPainter interface one can cast to (IIRC, from an IHTMLWindow2), which offers a Draw-method.
 There's also PrintPreview-interfaces which allow capturing of one (or more) pages.
 And of course you could try to work via API-calls directly on the hWnd of that Control, to copy its current Render-Output.

 Perhaps there's also some stuff hidden in all those IE-OleCmdID-EnumValues which might accomplish a Preview- or Capture-output.

 Olaf

## #15 SearchingDataOnly — Nov 8th, 2020, 05:42 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

The BindTo (as well as the two Navigate-Methods) have a second (optional)TimeOutSeconds-Parameter -
 by default sitting at 5 seconds.

 I guess your system is (shortly after booting) still quite busy with delayed loading of other base-stuff like services etc. -
 in addition to loading the quite large Chromium-Runtime, ... so the default of 5sec may be a bit too short -
 (I can imagine, that older magnetic disks might play a role, because I see no larger delay here on my test-OSes on SSDs).

 You could:
 - either try to increase that TimeOut-Param (to maybe 10sec or so)
 - or set it to Zero, to tell the methods to wait for their matching EventHandlers
 .. (WV_InitComplete or WV_NavigationCompleted respectively, to proceed with your own, now async Inits from within those Handlers)

Yes, after changing SecondsToWaitForInitComplete from 5 to 10, the problem was solved. Thank you very much.

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

You still planning on using the IE further???

I plan to gradually replace IE with WebView2 in the next period of time, but it will take some time. During this period, IE and WebView2 will exist in my system at the same time.

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Well, there's an IHTMLPainter interface one can cast to (IIRC, from an IHTMLWindow2), which offers a Draw-method.
 There's also PrintPreview-interfaces which allow capturing of one (or more) pages.

Ok. I will try this method.

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

And of course you could try to work via API-calls directly on the hWnd of that Control, to copy its current Render-Output.

Yes, I know this way. I can use SendMessage(...WM_PRINT...) to capture IE Preview.

 In addition, my current project is almost finished. In a few months, I will start to develop a scripting language and IDE similar to VB6 (of course, the development time is also calculated in years). If you could provide some good seeds (prototypes or suggestions), maybe I can develop a decent scripting language and IDE, and this IDE may be applied to your VB6-compatible compiler.

 (From dm)

## #16 Resurrected — Nov 10th, 2020, 09:32 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

This came as a surprise.

 Olaf, we can't thank you enough for your contribution.

 I haven't tested it yet but I am already thrilled.

 This is something I have been searching for for a long time.

## #17 saturnian — Nov 21st, 2020, 11:15 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf
 Fantastic job! I dreamed about it for many months and you did it, fabulous!

 Is there a way to force the localization of the cWebView2 component? On my computer, the sub menus are still in English despite an installation on a French Windows 10!

 The only way I have found to solve this problem is to destroy the EN-US.pak file in the Locales folder of the EdgeWebView installation and rename fr.pak to EN-US.pak! It works but it is not very clean!

 Do you have an idea ?

 Many Thanks

## #18 Schmidt — Nov 21st, 2020, 09:00 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

Is there a way to force the localization of the cWebView2 component?
 On my computer, the sub menus are still in English despite an installation on a French Windows 10!

Ah, yes... sorry.

 Forgot to reset my internal Lang-Handling "to normal" (it was still hardwired to "en" for my own tests) ...
 (fixed that just now, so the next RC6-version will start with the default-language of the current system again, regarding Popup-Menus and other localizations).

 But (regarding a workaround, applied to the version you currently have...) -
 there's always the last (optional) Param in the BindTo-method, which can "influence everything" (config-wise).

 BackGround:
 This last param (additionalBrowserArguments) can take up (and later represents) the commandline-args,
 which are passed to the Chromium-Process which gets started under the covers for every WebView.

 So, a simple change of the BindTo-line (passing an appropriate cmd-string):
 If WV.BindTo(picWV.hWnd, , , , "--lang=fr") = 0 Then ...
 or (for the chinese readers here, because it might not be obvious)
 If WV.BindTo(picWV.hWnd, , , , "--lang=zh-cn") = 0 Then ...
 ...should do it for you (until the next RC6-version, where - as said - the current locale will determine the default again).

 As for the list of available commandline-switches - it is quite a long one (but described for example here):
 <https://peter.sh/experiments/chromium-command-line-switches/> https://peter.sh/experiments/chromiu...line-switches/

 HTH

 Olaf

## #19 saturnian — Nov 22nd, 2020, 04:04 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

thank you so much, Olaf

## #20 saturnian — Nov 22nd, 2020, 02:18 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Another question, Olaf!
 Is there a way to intercept the URL when opening a new window after clicking on the "Open in new tab" submenu?
 This would allow programming the display of the new page in a new cWebView2

## #21 carl039 — Nov 22nd, 2020, 04:55 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thanks Olaf, this is really superb work. My VB6 app can finally be HTML5 compatible. *[img: Smilie]*

## #22 Schmidt — Nov 22nd, 2020, 08:44 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

Another question, Olaf!
 Is there a way to intercept the URL when opening a new window after clicking on the "Open in new tab" submenu?
 This would allow programming the display of the new page in a new cWebView2

Right, the old Param-amount in that Event was only allowing "suppression of the built-in NewWindow-Handling"
 (though did not give much info about "what to construct, in case one wants to create his own thing").

 So, this Event was enhanced in the new version 6.0.2, which I've just uploaded to the usual place...

 A Test-Handler for the original demo (showing the new incoming Event-Args) is the following one:

Code:

```
Private Sub WV_NewWindowRequested(ByVal IsUserInitiated As Long, IsHandled As Long, ByVal URI As String, NewWindowFeatures As RC6.cCollection)
  Debug.Print "URI: "; URI

  Dim i As Long 'Print out, what's sitting behind the NewWindowFeatures JSON-Collection
  For i = 0 To NewWindowFeatures.Count - 1
    Debug.Print NewWindowFeatures.KeyByIndex(i), NewWindowFeatures.ItemByIndex(i)
  Next

  IsHandled = 1 'prevent the WebView from "doing its own thing" about the new window
End Sub
```

HTH

 Olaf

## #23 saturnian — Nov 25th, 2020, 02:11 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thanks, Olaf! It works perfectly

## #24 xiaoyao — Nov 28th, 2020, 02:47 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

can you make a dll only support cWebView2?(Like 100kb?)
 how to call WebView2Loader.dll without rc6.dll?
 RC6.dll is too larger ,4MB.

## #25 xiaoyao — Nov 28th, 2020, 02:55 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

V8Engine.dll 1.34Mb, BY 2018
 stdcall OR CDECL,Quickjs also is good for test

 can test fastdb.dll for vb6?
 a memory db lib

## #26 saturnian — Nov 28th, 2020, 09:17 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf
 If I am using WebView2-BrowserControl on an MDIChild window, I have some issues.
 For example, the DocumentComplete event is not fired and the component is very slow
 What is the problem ? Do you have an idea ?

 François

## #27 Schmidt — Nov 28th, 2020, 09:22 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

... can you make a dll only support cWebView2?(Like 100kb?)

I could, but I won't...
 It doesn't make much sense to save 1 or 2 MB in deployment-size, when "the other needed part" -
 (the "evergreen WebView-runtime" from MS) is about 50MB already.

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

... how to call WebView2Loader.dll without rc6.dll?

Basically by writing your own binding:
 - first, by reworking the SDKs WebView2.**tlb **(making its COM-interfaces more "VB-friendly")
 - and then writing your own wrapper-class against the interfaces of this typelib (without exposing the typelib-types on the VB6-Class-interface)
 - the finally resulting VB6-Class then usually placed within your own VB6-AX-Dll, to avoid separate delivery of the WebView2.**tlb **

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

RC6.dll is too large ... 4MB.

The current deployment-size of the RC6 is about:
 - 3.4MB (in a *.zip archive ... with current inet-speeds, downloadable in about 0.3 seconds)
 - 2.5MB (in a *.7z archive ... with current inet-speeds, downloadable in about 0.2 seconds)

 And that makes it the smallest "DB- and GUI-including" Desktop-App-Framework on the planet ...
 (it is significantly smaller than e.g. GTK3, QT, Electron or "Avalonia+.NET-Core5")

 For most "non Hello-World-Apps" (which do something useful), you will have to deploy "a set of Dlls/OCXes alongside your Exe" anyways.

 So, when you develop and later deploy your App based on a dozen of smaller libs (Dlls and OCXes) - this will sum up to typically:
 - about 2-4MB in your deployment-zip (or perhaps 5-8MB, when you include Icon+ Image-Resources and maybe a Font or two).

 So no, I cannot see where the big "burden" from introducing the RC6 will come in
 (which would allow you, to replace most of your other COM-dependencies with a single one in your App).

 Olaf

## #28 Schmidt — Nov 28th, 2020, 09:34 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

Hello Olaf
 If I am using WebView2-BrowserControl on an MDIChild window, I have some issues.
 For example, the DocumentComplete event is not fired and the component is very slow
 What is the problem ? Do you have an idea ?

What OS is that scenario running on?
 And could you post a little "mockup.zip" which includes the basic-setup (with the MDI-ChildWindows and the Bindings on them) -
 and which would allow me to reproduce what you've seen?

 Another related question would be...
 If you change from "MDIChildWindow-hosting" to "normal Forms + PictureBoxes" (maybe behind a "tabbed interface") -
 do you see differences in behaviour (when running against the same http-server, using the same HTML5-inputs)?

 If yes, then it might be something related to the VB6 MDIChildWindow-class (and its slightly "off-standard" message-pumping/handling underneath)...

 Olaf

## #29 Schmidt — Nov 28th, 2020, 09:52 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

V8Engine.dll 1.34M, BY 2018
 stdcall OR CDECL,Quickjs also is good for test

As already stated in my reply to dreammanor (dm) here:
 You can already make use of "V8" via the WebView2-Binding
 (which will then include a proper "window-object" for your V8-scripts already).

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

can test fastdb.dll for vb6?
 a memory db lib

I've already done such a test-comparison (in a former reply to you).
 And the result was, that the RC5/RC6-wrapper for SQLite performs better than "fastdb".

 In-Memory-DB scenarios are possible via SQLite as well:
 - either via the "normal" cConnection-SQLite-wrapper Class in "::memory::"-Mode
 - or via the cMemDB-class (which offers a few more convenience-functions on top of cConnection)

 Olaf

## #30 saturnian — Nov 28th, 2020, 10:31 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I found !
 I have misused the properties of cWebview!

 'WV.AreDefaultContextMenusEnabled = True
 'WV.AreDefaultScriptDialogsEnabled = True
 'WV.AreDevToolsEnabled = True
 'WV.AreHostObjectsAllowed = True
 'WV.IsScriptEnabled = True
 'WV.IsWebMessageEnabled = True

 WV.IsStatusBarEnabled = True

 If I uncomment this whole block of code, it doesn't work! Otherwise everything works fine (just with WV.IsStatusBarEnabled = True )

## #31 Schmidt — Nov 28th, 2020, 12:32 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

I found !
 I have misused the properties of cWebview!

 'WV.AreDefaultContextMenusEnabled = True
 'WV.AreDefaultScriptDialogsEnabled = True
 'WV.AreDevToolsEnabled = True
 'WV.AreHostObjectsAllowed = True
 'WV.IsScriptEnabled = True
 'WV.IsWebMessageEnabled = True

 WV.IsStatusBarEnabled = True

 If I uncomment this whole block of code, it doesn't work! Otherwise everything works fine (just with WV.IsStatusBarEnabled = True )

That's a bit surprising, because all of these Props are (by default, after initialization) already "On" (at C++ Boolean-Value = 1).

 What you should keep in mind though is, that I've implemented these Boolean-Props "As Long"
 (to directly map between the C++ Boolean prop-calls and the Wrapper-Props).

 So, what gets passed along (into the C++ iplementation of the WebView) in your example above,
 are (due to VBs automatic Type-coercing) "Long-Values which are -1" (coerced from VB6-True).

 So these "-1" Values you pass (in a 32Bit-slot, having all bits at 1) might confuse the C++-code (although, normally this should not be an issue).

 That said, ... testing.... and arrrghhh...
 Whilst checking this in the little Demo-App of the opener thread (placing your calls after the BindCall),
 you are right indeed - the True Values (transported as -1 Longs) confuse the WebView2 here as well... (every Prop-setting-line being active)

 But good news... I have no issues when changing "all the True's" to the C++ Value 1:

Code:

```
Private Sub Form_Load()
  ...

  If WV.BindTo(picWV.hWnd) = 0 Then MsgBox "couldn't initialize WebView-Binding": Exit Sub
  Debug.Print WV.GetMostRecentInstallPath

  WV.AreDefaultContextMenusEnabled = 1
  WV.AreDefaultScriptDialogsEnabled = 1
  WV.AreDevToolsEnabled = 1
  WV.AreHostObjectsAllowed = 1
  WV.IsScriptEnabled = 1
  WV.IsWebMessageEnabled = 1
  WV.IsStatusBarEnabled = 1

  LocalWebViewInit 'initialize the WebView for local usage here in our Form
End Sub
```

Ok then, please expect the next RC6 version to have proper VB6-Boolean types on all the relevant Config-Props.
 (which I will then take care of internally, mapping them to the correct C++ Boolean-Type).

 So, thanks for your testing, which made this behaviour apparent in this still "early stage"
 (where the RC6-interfaces for its new Class-Additions are not yet "finalized").

 Olaf

## #32 xiaoyao — Nov 29th, 2020, 12:59 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

by vfb maybe no need rc6.dll,WebView2Loader.dll
 it'S like vc++

## #33 saturnian — Dec 6th, 2020, 10:30 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

So, thanks for your testing, which made this behaviour apparent in this still "early stage"
 (where the RC6-interfaces for its new Class-Additions are not yet "finalized").

 Olaf

This is nothing compared to the fabulous work you do for the community !
 Thanks again, Olaf

 To test, I wrote a multi-tab browser, which begins to work well.
 I realized that it was necessary to avoid giving the focus to another control, in the NavigationCompleted event. If you want to modify a display based on an error number, for example, it is better to do so by activating a timer which will trigger the display late.

 Other thing : the URLs "chrome://flags", "chrome://history", "chrome://downloads",... work but do not trigger the DocumentComplete event, which sometimes causes the Webview component to lag for several tens of seconds! I don't know why !

 Sincerely

 François

## #34 Schmidt — Dec 9th, 2020, 02:22 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

Other thing : the URLs "chrome://flags", "chrome://history", "chrome://downloads",... work but do not trigger the DocumentComplete event, which sometimes causes the Webview component to lag for several tens of seconds! I don't know why !

Perhaps these "internal URLs" do not really "manifest" themselves in a true document...
 Did you try to use the NavigationCompleted-event for these URLs instead?

 Olaf

## #35 SearchingDataOnly — Dec 16th, 2020, 10:08 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

(From dm)

 I've bound WebView to cwMyWidget and it works very well. Now I have two questions:
 (1) How to add my own right-click menu to WebView?
 (2) Is it possible to bind IE-WebBrowser-Control (Shell.Explorer.2) to cWidgetBase or cWidgetRoot? If this is possible, I can keep both WebView2 and IE-WebBrowser in my system. Thanks.

 **Edit:**
 The first question has been solved:
 Step1: WV.AreDefaultContextMenuEnabled = False
 Step2: WV_UserContextMenu Event

## #36 SearchingDataOnly — Dec 16th, 2020, 10:50 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The third question:
 (3) When using WV.Navigate to open an html file, everything is OK. The external css files and js files in the html file can be accessed normally. However, when using WV.NavigateToString to open an html-string, the external css files or js files contained in the html-string will not be accessible, and DevTools prompts "Not allowed to load local resource...".

 I'd like to know if there is a way to solve this problem (for example, can this problem be solved by cWebServer?). In my software, all Html operations are done through html-strings instead of html-files. In other words, I need to use WV.NavigateToString a lot instead of WV.Navigate a file. (Note: There is no such problem in IE-WebBrowser)

 In addition, after I used the parameter "--disable-web-security --allow-file-access-from-files --allow-file-access", DevTools prompts "Failed to load resource: net::ERR_UNKNOWN_URL_SCHEME"

 **Edit:**
 The problem discussed in the following thread is similar to mine:
 <https://github.com/MicrosoftEdge/WebView2Feedback/issues/642> https://github.com/MicrosoftEdge/Web...ack/issues/642

 The poster seems to have solved the problem through SetVirtualHostNameToFolderMapping, but I have not seen the SetVirtualHostNameToFolderMapping API in RC6.WebView.

 **Edit 2:**
 Two other links related to this issue:

 API Spec for NavigateWithWebResourceRequest
 <https://github.com/MicrosoftEdge/WebView2Feedback/pull/483/files/4ddd96fa846142b535af64eabe9267942e836977> https://github.com/MicrosoftEdge/Web...9267942e836977

 WebView.NavigateToLocalStreamUri(Uri, IUriToStreamResolver) Method
 <https://docs.microsoft.com/en-us/uwp/api/windows.ui.xaml.controls.webview.navigatetolocalstreamuri?view=winrt-19041#Windows_UI_Xaml_Controls_WebView_NavigateToLocalStreamUri_Windows_Foundation_Uri_Windows_Web_IUriToStreamResolver_> https://docs.microsoft.com/en-us/uwp...treamResolver_

## #37 saturnian — Dec 17th, 2020, 06:27 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Did you try to use the NavigationCompleted-event for these URLs instead?

 Olaf

Hello Olaf

 I, of course, tried to use the NavigationCompleted event. But trying to give focus to another object in this event freezes the cWebView component for several minutes, often. I solved this by activating a timer in the NavigationCompleted event which itself, a second after, update the other objects (The form caption with the tltle of the web page, for example)

## #38 saturnian — Dec 17th, 2020, 06:33 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf

 is it possible to expose the "NavigationStarting" event in cWebView2? It seems to me that this is the only event in Microsoft's WebView2 control that is not present in cwebView2!

 Kindly
 François

## #39 SearchingDataOnly — Dec 22nd, 2020, 02:27 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The fourth question:
 (4) I need to place multiple WebViews on my Form or WidgetForm. I bind a WebView to a cwWidget through "WV.BindTo(W.Root.hWnd, 10)".
 The question is, when this cwWidget is removed, how to remove the bound WebView at the same time?

## #40 Thierry76 — Feb 3rd, 2021, 02:40 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thanks a lot Olaf !
 I try a little your code last sunday and it's work well ....
 I have one question, did somebody find a way to install an Edge extension by code on the WebView2 ? Is it possible ? I beggin in this sort of developement so excuse me if my question is stupid...

 PS I dream of your webviev2 with diagrams.net (drawio-desktop)....

## #41 Schmidt — Feb 4th, 2021, 02:25 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Thierry76** *[img: View Post]*

PS I dream of your webviev2 with diagrams.net (drawio-desktop)....

Why not just load a js-based charting-lib into the WebView (instead of a plugin)?

 There's a dozen or more quite good ones out there -
 the one we've decided to use is dc.js (because it's open - and based on the svg+canvas-lib d3.js, which we had in use already):

 Here's a quite impressive example (which supports zooming and range-shifting on an interdependent, complex dataset):
 <https://dc-js.github.io/dc.js/> https://dc-js.github.io/dc.js/

 And here a turorial for it (normally, you'll only need a dozen of js-config-lines, to visualize a certain chart-type).
 <https://www.tutorialspoint.com/dcjs/index.htm> https://www.tutorialspoint.com/dcjs/index.htm

 If you're more interested in org- and flow-charts... simple ones can be created directly in code via d3.js -
 but there's also libs out there which specialize in that kind of thing ...
 (though I have not needed any org- and flowchart functionality so far, so cannot give any useful hints).

 Olaf

## #42 xiaoyao — Feb 7th, 2021, 12:18 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **wqweto** *[img: View Post]*

@xiaoyao: Does VFB have bindings for the WebView2 built-in?

 What's the point spamming every thread with links to it?

 @admins: Do cleanup my reply too, please.

 cheers,
 </wqw>

YES VFB WITH MINIBLINK webbrowser,chrome core ,node.dll

## #43 xiaoyao — Feb 7th, 2021, 12:28 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

I could, but I won't...
 It doesn't make much sense to save 1 or 2 MB in deployment-size, when "the other needed part" -
 (the "evergreen WebView-runtime" from MS) is about 50MB already.

 Basically by writing your own binding:
 - first, by reworking the SDKs WebView2.**tlb **(making its COM-interfaces more "VB-friendly")
 - and then writing your own wrapper-class against the interfaces of this typelib (without exposing the typelib-types on the VB6-Class-interface)
 - the finally resulting VB6-Class then usually placed within your own VB6-AX-Dll, to avoid separate delivery of the WebView2.**tlb **

 The current deployment-size of the RC6 is about:
 - 3.4MB (in a *.zip archive ... with current inet-speeds, downloadable in about 0.3 seconds)
 - 2.5MB (in a *.7z archive ... with current inet-speeds, downloadable in about 0.2 seconds)

 And that makes it the smallest "DB- and GUI-including" Desktop-App-Framework on the planet ...
 (it is significantly smaller than e.g. GTK3, QT, Electron or "Avalonia+.NET-Core5")

 For most "non Hello-World-Apps" (which do something useful), you will have to deploy "a set of Dlls/OCXes alongside your Exe" anyways.

 So, when you develop and later deploy your App based on a dozen of smaller libs (Dlls and OCXes) - this will sum up to typically:
 - about 2-4MB in your deployment-zip (or perhaps 5-8MB, when you include Icon+ Image-Resources and maybe a Font or two).

 So no, I cannot see where the big "burden" from introducing the RC6 will come in
 (which would allow you, to replace most of your other COM-dependencies with a single one in your App).

 Olaf

more times,i no need rc6.dll.
 but i need Chromium or edge ,or miniblink (by Chromium),so if only one dll vbedgeCom.dll or stdcall_edge.dll,
 i think it's best
 if it's make ocx control,i think it's easy for vb user,excel vba userform
 if it's support x64 like a stdcall dll,it's also can use in vfb ide FOR X64 EXE (VISUAL freebasic)

## #44 Schmidt — Feb 7th, 2021, 03:39 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

...more times,i no need rc6.dll.

Then this thread is not for suitable for you.
 <shrug/>

 What's discussed here, is a Binding to the **EverGreen**-Runtime of the MS-Edge-Chromium-WebControl.

 EverGreen meaning, that this Runtime has to be installed **only once** on a given Client-Machine.
 After that, the regular Windows-Edge-Updates will update also this WebControls Chromium-Runtime without any further User-intervention.

 So, the Client-Machine is guaranteed, to always use a WebView-Control with the latest Security- and Bug-Fixes.

 This is not the case for e.g. "miniblink" - which (have just checked in their Repo) is already 2 years out of date.

 Olaf

## #45 rameshpatel9 — Feb 21st, 2021, 02:06 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,

 Is there a way to implement late-binding for WebView2? This is to handle a scenario where user has not yet installed pre-requisites (WebView2 Runtime and RC6).

## #46 Schmidt — Feb 22nd, 2021, 08:02 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **rameshpatel9** *[img: View Post]*

Is there a way to implement late-binding for WebView2?
 This is to handle a scenario where user has not yet installed pre-requisites (WebView2 Runtime and RC6).

To make use of the WebView2-EdgeChromium-BrowserControl,
 you have to "bind it to a normal hWnd" (in the Demo, this is provided by a VB6.PictureBox).

 And that Binding is initiated via an RC6-Class (cWebView).

 So, the RC6 is a pre-requisite to make all that work.

 Not sure, what you mean with "late-binding" - but in case you mean "without running a prior setup, or ensuring registry-entries" -
 that's possible when you:
 - ship the 5 RC6-Dlls in a \Bin\ Subfolder of your deployment-zip
 - then instantiate New_c regfree within your App (via a declared DirectCOM.dll GetInstance or GetInstanceEx-call)
 - and then initiate the Binding by a Class-instance, derived via New_c.WebView

 The only other pre-requisite is, that the Users have installed the "EverGreen-runtime" from the MS-link I gave in the First posting here.

 HTH

 Olaf

## #47 oskarrr — Mar 8th, 2021, 09:51 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello!

 I am using WebView2 during 6 months. It was working perfectly.

 But now, stop work without reason!

 The strange thing is that: From compiler VB6 it works perfectly.
 But if I generate .exe file, don´t work: "Can not initialize webview2"

 I can not find where is the problem.
 Any idea?

 I reinstall component, register dll again...

 Thank you!

## #48 jpbro — Mar 8th, 2021, 12:23 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Unfortunately, it looks like MS might have broken something: <https://www.vbforums.com/showthread.php?890812-RC6-and-WebView2-Error-in-compiled-exe> https://www.vbforums.com/showthread....n-compiled-exe.

## #49 oskarrr — Mar 8th, 2021, 04:21 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jpbro** *[img: View Post]*

Unfortunately, it looks like MS might have broken something: <https://www.vbforums.com/showthread.php?890812-RC6-and-WebView2-Error-in-compiled-exe> https://www.vbforums.com/showthread....n-compiled-exe.

wow, I was getting crazy. Fortunatly, I have followed the instructions of "saturnian" and now works perfectly.

 Thank you so much!

## #50 WZabel — Mar 11th, 2021, 05:46 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

First of all, many thanks to Olaf for this great work.

 Meanwhile a newer version od WebView2 has been released, 89.0.774.50 is working with IDE and compiled programs.

 To avoid such problems in future i tried to use a fixed version, with no success. *[img: Frown]*

Code:

```
BindTo(HosthWnd As Long, [SecondsToWaitForInitComplete As Double = 8], [browserInstallPath As String], [userDataFolder As String], [additionalBrowserArguments As String]) As Long
```

is not binding when setting the second argument.

 Has anyone tested this before?

 Robert

## #51 xiaoyao — Mar 11th, 2021, 05:52 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

how to call by vc++ for show edge like a ocx on wpf(window,or put to hwnd)

 I want call by vfb(visual freebasic ide),I don't want use ocx or com dll

## #52 Schmidt — Mar 11th, 2021, 06:29 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **WZabel** *[img: View Post]*

Meanwhile a newer version od WebView2 has been released, 89.0.774.50 is working with IDE and compiled programs.

Yep, I've described that here: <https://www.vbforums.com/showthread.php?890812-RC6-and-WebView2-Error-in-compiled-exe&p=5513856&viewfull=1#post5513856> https://www.vbforums.com/showthread....=1#post5513856

>

>

 *[img: Quote]* Originally Posted by **WZabel** *[img: View Post]*

To avoid such problems in future i tried to use a fixed version, with no success. *[img: Frown]*

I've made a few comments about that in the above link as well...

>

>

 *[img: Quote]* Originally Posted by **WZabel** *[img: View Post]*

Code:

```
BindTo(HosthWnd As Long, [SecondsToWaitForInitComplete As Double = 8], [browserInstallPath As String], [userDataFolder As String], [additionalBrowserArguments As String]) As Long
```

is not binding when setting the second argument.

The second argument is the timeout you plan to wait for the Init-Process to finish...

 If you want to set the EdgeInstall-Path, this has to be done in the 3rd-Parameter.

 Here is an example, how I was able to load directly from a fixed version (in my App.Path):

Code:

```
Private Sub Form_Load()
  Visible = True '<- it's important, that the hosting TopLevel-Form is visible...

  Set WV = New_c.WebView2 'create the instance

  Dim P As String
'      P = WV.GetMostRecentInstallPath() '<- this is what I use underneath the BindTo-call, when you don't give a specific path
      P = App.Path & "\Microsoft.WebView2.FixedVersionRuntime.89.0.774.50.x86"
  Debug.Print P

  If WV.BindTo(picWV.hWnd, 5, P, App.Path) = 0 Then MsgBox "couldn't initialize WebView-Binding": Exit Sub
'  If WV.BindTo(picWV.hWnd) = 0 Then MsgBox "couldn't initialize WebView-Binding": Exit Sub

  '...
```

HTH

 Olaf

## #53 Schmidt — Mar 11th, 2021, 06:37 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

... I don't want use ocx or com dll

Then you are out of luck - because the WebView2Loader.dll from MS *is* a COM-dll.
 (providing it's COM-Objects over "flat-API-entry-call" - and later on "building on that" via a matching typelib)

 And as always, your postings do not make much sense
 (WPF belongs to the .NET-universe - and VFB is "Class-less, COM-less" flat-APIs only)

 The only IDE which can "talk" to the WebView2Loader-Objects directly
 (without any "extra-marshalling" or "COM-call-support-modules"), is the VB6-IDE.

 And we are here in a Sub-Forum where VB6 is discussed...
 Please stop spamming this Forum with your nonsense.

 Olaf

## #54 xiaoyao — Mar 14th, 2021, 10:26 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **wqweto** *[img: View Post]*

@xiaoyao: Does VFB have bindings for the WebView2 built-in?

 What's the point spamming every thread with links to it?

 cheers,
 </wqw>

I didn't fully express my problem. I'm sorry.
 vfb support miniblink,chrome core only one DLL.
 edge like google beowser.For example, running speed and Lesser memory, there are many advantages.
 So I would like to ask VfB how to call the edge, the main support for com is not perfect, so is there anyway only a DLL can be called?(without rc6.dll)

 Previously, VfB did not support the creation of com DLLs, but now it can create the built-in TLB information. You know a lot about com OCX. It would be more convenient if you could increase the build to OCX.

## #55 Schmidt — Mar 14th, 2021, 03:19 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

vfb support miniblink,chrome core only one DLL.
 ... running speed and Lesser memory, there are many advantages.

That is not true, because in the end there's always separate Chromium-Processes which will be spawned.
 (and those ChildProcesses will have the basically same performance and mem-consumption).

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

So I would like to ask VfB how to call the edge...

You contradict yourself.

 Why would you now suddenly need an Edge-Binding for FreeBasic, when "miniblink" is so much better (in your opinion).

 But as already said, stop spamming these threads with unrelated things like FreeBasic
 (I consider FB a great tool - but it has no place in this VB6-related Forum-Thread).

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

You know a lot about com OCX.
 It would be more convenient if you could increase the build to OCX.

I don't really know, what exactly you are talking about here - but for your information,
 a COM-library which exists as a Dll, is far more versatile (and much easier to use) -
 compared to a COM-lib which offers "visual COM-Objects" behind an OCX-incarnation.

 That's with regard to "instancing"...
 A "normal COM-Object-instance" is much easier to load from a COM-Dll,
 compared to the quite messed up "siting-processes" which are needed in case of a "visual COMObj-instance from an OCX".

 Olaf

## #56 xiaoyao — Mar 14th, 2021, 08:48 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

vb6 Slowly be eliminated,so I study vfb,it'S like vc++.
 Microsoft. Net framework, but also weakened the COM, OCX control package.So I want to ask if there is a way to call edge by Stdcall?

 WebView2Loader.dll ,why Double-clicking will not prompt you that you have successfully registered.

 The main reason is that some computer permissions are not enough and do not support the registration com DLL.

 You may be able to call com objects using just the windows API. So I want to learn how to create a simple WebView 2 page,How many windows APIs do you need to call ,for load webview2?
 If rc6. Dll exports a stdcall apiThat would be too convenient.
 dim web as long
 web =loadWebviewTo(picture1.hwnd)
 webviewOpenUrl(web,"http://*")

 In fact, There is an open source browser miniblink" like a activex control OCX ( chromium,only need one dll) is using this principle to operate the Google Chrome Kernel,This is a reduced version of the feature, and the version is relatively old.
 edge with chromium ,The function is relatively complete,It's better than the miniblink. "A lot more powerful.".

## #57 byomi1 — May 4th, 2021, 03:16 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,

 a shameless semi-offtopic question related VBA / MSAccess:

 Any chance I can get Webview2 working in MS Access / VBA???

 Greets from Germany

## #58 xiaoyao — May 4th, 2021, 03:41 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **byomi1** *[img: View Post]*

Hi Olaf,

 a shameless semi-offtopic question related VBA / MSAccess:

 Any chance I can get Webview2 working in MS Access / VBA???

 Greets from Germany

Edge-Binding,Microsoft's new version of WIN10 may have a built-in Google kernel EDGE browser. In addition, the COM interface provided by Microsoft has been generally used for several years and is relatively stable and perfect.
 You can also use VB6 to develop an OCX control （put WebView2-Binding on usercontrol）, and then load the user-defined control to the EXCEL, ACCESS VBA form

## #59 SteveMB — Aug 2nd, 2021, 05:10 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi,

 I've been trying to get the demo from the first post to run. I've downloaded the latest RC6 version from RichClient-lib (Version 6.0.8) and the evergreen webview2 runtime installer from the given microsoft link. The webview2 runtime installed successfully, as did RC6 (using the VBscripts in the zip package). As recommended, all of the dlls are in the same folder.

 When I try to compile the WebView2Demo, I'm getting the following error message when it reaches the event handlers (with parameters) in fMain.frm

 Compile error: Procedure declaration does not match description of event or procedure having the same name

 Is this demo now obsolete (and is there a newer version somewhere)? Or have Microsoft changed things again and broken this functionality completely?

## #60 jpbro — Aug 2nd, 2021, 10:36 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **SteveMB** *[img: View Post]*

When I try to compile the WebView2Demo, I'm getting the following error message when it reaches the event handlers (with parameters) in fMain.frm

 Compile error: Procedure declaration does not match description of event or procedure having the same name

It looks like some of the RC6 cWebView2 procedure declarations have changed. If you modify the declarations for the following 2 events, it should work:

Code:

```
Private Sub WV_AcceleratorKeyPressed(ByVal KeyState As eWebView2AccKeyState, ByVal IsExtendedKey As Boolean, ByVal WasKeyDown As Boolean, ByVal IsKeyReleased As Boolean, ByVal IsMenuKeyDown As Boolean, ByVal RepeatCount As Long, ByVal ScanCode As Long, IsHandled As Boolean)

Private Sub WV_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
```

 Some "Longs" were changed to be more sensible "Booleans".

## #61 SteveMB — Aug 2nd, 2021, 03:15 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jpbro** *[img: View Post]*

It looks like some of the RC6 cWebView2 procedure declarations have changed. If you modify the declarations for the following 2 events, it should work:

Code:

```
Private Sub WV_AcceleratorKeyPressed(ByVal KeyState As eWebView2AccKeyState, ByVal IsExtendedKey As Boolean, ByVal WasKeyDown As Boolean, ByVal IsKeyReleased As Boolean, ByVal IsMenuKeyDown As Boolean, ByVal RepeatCount As Long, ByVal ScanCode As Long, IsHandled As Boolean)

Private Sub WV_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
```

 Some "Longs" were changed to be more sensible "Booleans".

Thanks jpbro, nice to see it was something relatively simple to fix. I've got the demo compiling and running now. Now I'll have to work out how best to build it into my application and see how it runs there.

## #62 SearchingDataOnly — Aug 2nd, 2021, 07:52 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **SteveMB** *[img: View Post]*

Thanks jpbro, nice to see it was something relatively simple to fix. I've got the demo compiling and running now. Now I'll have to work out how best to build it into my application and see how it runs there.

Maybe the following example is useful to you:
 <https://www.vbforums.com/showthread.php?892809-RC6-WebView2-Demo> https://www.vbforums.com/showthread....-WebView2-Demo

## #63 jackr — Aug 30th, 2021, 03:27 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,

 I'm trying to install Webview2Demo. I followed the instructions but I get the error:"Procedure declaration does not match description of event or procedure having the same name". Any ideas why. Thank you very much!

 -Jack

## #64 jackr — Aug 30th, 2021, 04:06 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

one question:

 using the example and navigating to "google.com", I can't seem to be able to get the "document.getElementById('btn1').innerHTML" working using WV.jsProp. Any ideas? basically, I need the html of the entire page I just navigated to. I assume I can do so using "body.innerHTML"

## #65 SearchingDataOnly — Aug 30th, 2021, 06:32 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jackr** *[img: View Post]*

Hi Olaf,

 I'm trying to install Webview2Demo. I followed the instructions but I get the error:"Procedure declaration does not match description of event or procedure having the same name". Any ideas why. Thank you very much!

 -Jack

The answer you want is in post#62

## #66 jackr — Aug 31st, 2021, 08:42 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

The answer you want is in post#62

thanks that works!

## #67 jackr — Aug 31st, 2021, 11:00 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Is there a way to invoke a js click such as
 WV.jsRun("document.getElementById('Somebutton')".click()")
 I was unable to invoke a click with that

 thanks for your help!

## #68 Schmidt — Aug 31st, 2021, 12:21 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jackr** *[img: View Post]*

Is there a way to invoke a js click such as
 WV.jsRun("document.getElementById('Somebutton')".click()")

WV.jsRun is thought for simple (for the most part, self-defined) functions...

 For more complex "actions" like yours above, try it with:
 WV.ExecuteScript "document.getElementById('Somebutton').click()"
 instead.

 HTH

 Olaf

## #69 jackr — Aug 31st, 2021, 01:47 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

That helps, thanks again!!

## #70 xiaoyao — Aug 31st, 2021, 08:41 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

what's MicrosoftEdgeWebView2Runtime?
 Is the development version of the interface open?

 How to operate the EDGE that is installed or that comes with WIN10, can I control the CHROME kernel without installing any RUNTIME? Maybe only Google Control Protocol, based on HTTP port, like WEBDRIVER?

## #71 xiaoyao — Aug 31st, 2021, 08:46 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The CHROME kernel environment, there are three ways to get:

 Install the EDGE (Chromium) of the development version, the stable version of the EDGE does not currently support WebView2 controls.
 Install independent WebView2 Runtime, which can download and upgrade an EDGE Chromium kernel separately
 Embedding Edge Chromium kernel

 I don't know if there's a way. Do not need to install, only need to extract the file, and Have the least hard disk space

## #72 xiaoyao — Aug 31st, 2021, 08:59 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Edge Webview2 Automatic installation, detect the installed version-VBForums
 <https://www.vbforums.com/showthread.php?893196-Edge-Webview2-Automatic-installation-detect-the-installed-version> https://www.vbforums.com/showthread....talled-version

 How to use code-behind automatic installation MicrosoftEdgeWebView2RuntimeInstallerX86？

 <https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/0e69d99b-247e-4a2c-a6ca-a7c66abd4cda/MicrosoftEdgeWebView2RuntimeInstallerX86.exe> https://msedge.sf.dl.delivery.mp.mic...stallerX86.exe

 Which URL can use the code to automatically download the latest version? Plus automatic installation is very convenient

 【install used 32 seconds, occupying 444mb of hard disk space】
 use miniblink.DLL or cef.DLL(10-50MB)

 For the chromium kernel, only one DLL is needed. Use this mini-small WEBKIT-related kernel as much as possible. If some web pages are too complicated and require the latest HTML5 technology to display, then use Microsoft's EDGE runtime library. 450Mb is more than 10-50mb, which takes up too much hard disk space.

## #73 xiaoyao — Aug 31st, 2021, 09:30 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Code:

```
Function CheckSetupOk() As Boolean
    'It takes 10 seconds to check whether the edge runtime component is installed successfully. Is there a faster method?
    Dim WV As cWebView2
    Set WV = New_c.WebView2
    CheckSetupOk = WV.BindTo(Me.hWnd) <> 0
    Set WV = Nothing
End Function

Function DonwSetupTool() As Boolean
Dim URL As String
URL = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
'xmlhttp download***
'save as :Edge_Webview2RunTime.exe

Dim Size1 As Long
Size1 = FileLen(App.Path & "\Edge_Webview2RunTime.exe")
DonwSetupTool = Size1 > 1024 ^ 2

end function
```

install used 32 seconds, occupying 444mb of hard disk space
 I don’t know if there is a silent installation parameter to prevent him from displaying the download and installation interface

 **MicrosoftEdgeWebview2Setup.exe /silent /install **

 check Registry information：
 Microsoft Edge WebView2 Runtim(92.0.902.8)
 HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}
 get string value(PV and name)

## #74 softv — Sep 4th, 2021, 09:39 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

I could, but I won't...
 It doesn't make much sense to save 1 or 2 MB in deployment-size, when "the other needed part" -
 (the "evergreen WebView-runtime" from MS) is about 50MB already.
 ... .. .
 Olaf

Is there a necessity that the end-user systems (where my application is installed) also should have "evergreen WebView-runtime" installed in them (if not already installed)?

 Kind regards.

## #75 xiaoyao — Sep 4th, 2021, 11:01 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

You can keep it updated automatically. It's also very convenient. Ie is very rubbish, the edge of the Microsoft's last Straw, he will certainly operate very well.

 But if you want to take up less memory, take up less hard disk space, install a fixed old version, this is a better choice.

## #76 Schmidt — Sep 4th, 2021, 11:21 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **softv** *[img: View Post]*

Is there a necessity that the end-user systems (where my application is installed) also should have "evergreen WebView-runtime" installed in them (if not already installed)?

Yes, of course - but on each of the target-systems, this "base-runtime-installation" will only have to be done once
 (as the "evergreen" suggests, this chromium-runtime will from this point on, be updated by the system).

 As for detection, whether a given target-system already contains this runtime,
 you can simply check the output of the GetMostRecentInstallPath-method:
 strInstallPath = WV.GetMostRecentInstallPath

 If it gives back an empty string, then you need to inform the user,
 to download the little 1.7MB base-installer from the MS-webpage via this (apparently constant) Link -
 (as already given in the Opener-Posting of this thread):
 <https://go.microsoft.com/fwlink/p/?LinkId=2124703> https://go.microsoft.com/fwlink/p/?LinkId=2124703

 If you ask your user beforehand like this:

Code:

```
  If Len(WV.GetMostRecentInstallPath) = 0 Then 'Edge-Chromium-Webview is missing on this system
     If MsgBox("The needed MS-WebView2-runtime is missing," & vbLf & "Do you want to download the installer now?", vbYesNo) = vbYes Then
        New_c.FSO.ShellExecute "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
     End If
  End If
```

The little 1.7MB MS-Executable from that download above, will then determine which system the User actually runs (Win7, Win8, Win10) -
 and then download and install the larger (about 50MB-80MB) and properly matching evergreen-runtime which will work best on the given system.

 HTH

 Olaf

## #77 Schmidt — Sep 4th, 2021, 11:24 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

But if you want to take up less memory, take up less hard disk space, install a fixed old version, this is a better choice.

Please stop giving misleading advice here in this thread.

 Everything you've posted so far here (including your posts #72, #73, #74, #75) -
 contains only your usual nonsensical garbage.

 Olaf

## #78 xiaoyao — Sep 4th, 2021, 12:56 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Because at first I didn't know if. Net directly supported the standard version of the edge.
 Just like ie, I thought the system came with WebBrowser controls.
 After the win10, win11, the system recommended by default to install the edge browser software, using the Google kernel, if you can bring webview2 control is convenient, otherwise you need to install 500MB runtime.

 I just like automatic installation and deployment.The original detection of whether the edge runtime is installed, it takes more than 10 seconds, so later found a way to find the registry.
 I spent more than a day on it. I just hope others can avoid detours.
 Then I stumbled on the silent installation method.

## #79 xiaoyao — Sep 4th, 2021, 01:46 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thank you for your contribution to the highly encapsulated Google kernel.
 I just spent more than an hour reading everyone's replies.Many people do not know the problem they solved, but did not leave detailed code.
 Is there a way to open the privacy page?
 For example, create two webview objects in a form.
 But the cookie between them is quarantine each other, is there such a setting?
 If you create objects in multiple threads and use multiple webview controls in the same window, can you improve the running speed?

 IE WebBrowser may run in the STA mode. If you create it in multi-thread, the webpage will not show, even if the creation is successful, the running speed is also very slow, and it is actually a message queue processing, eventually is running in the same thread. in.

## #80 xiaoyao — Sep 4th, 2021, 01:57 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **rameshpatel9** *[img: View Post]*

Hi Olaf,

 Is there a way to implement late-binding for WebView2? This is to handle a scenario where user has not yet installed pre-requisites (WebView2 Runtime and RC6).

After my test. The new computer is not installed with WebView Runtime, nor will it be wrong. This is a creating an object, I will be snatched 10 seconds, then the prompt fails. In this case, the automatic download of the runtime installer will be successfully created after the installation is complete.

 There is also a problem, rc6.dll. Are there any plans to launch a 64-bit version?
 In this way, VBA, 64-bit office, VB. Net and other IDE or development tools can also use it.
 Twinbasic, and many other IDE alternatives to VB6 are under development, many of which can generate 64-bit exeThe biggest disadvantage of Microsoft's VB6 is that it does not support 64-bit, but VBA has a 64-bit version

 Are there any plans to develop a simple webdriver like interface?
 It just uses HTTP or websocket to control a properly installed Google Chrome or Edge browser.

 You do not need to install a development version of edge or a webview runtime.

 Almost eighty or ninety percent of computers have Google Chrome installed.
 There is an open source project based on the COM interface of Google webview driver. But it hasn't been updated for many years.

 A good technician can compile it to 64 bits.And add some of the latest version of the function, because there are some functions have been wrong or outdated.

## #81 xiaoyao — Sep 4th, 2021, 02:19 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

vb6 webdriver,support vba,call edge chrome,
 msedgedriver.exe

 Without systematic teaching materials, many people will only use simple open web pages.

 I've tested this before, getting all the tabs, switching between any of them, and it's very time consuming to debug.

 So if you can have a master sub-installed easier to use, many people will not need to learn to use.

 JS, database, network programming, interface beautification, Google browser control, these areas have great market potential in the future.

 <https://blog.csdn.net/qq_24499417/article/details/105602339> https://blog.csdn.net/qq_24499417/ar...ails/105602339
 <https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/>
 https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/

## #82 xiaoyao — Sep 4th, 2021, 02:45 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Microsoft WIN11 preinstalled messaging APP teams 2.0 experience: Built on Edge Webview 2 2021-08-18

 Just saw the news, the future win11 comes with a webview2 runtime.
 There is no need to download and install it manually in the future, so it is much more convenient.

 If we study webdriver more thoroughly, what browser is automatically detected on the client? Automatically download the driver, automatically open the web page and automatically fill in the form.

 Applications such as Microsoft Office have begun to deploy WebView2 Runtime in their applications. WebView 2 Runtime has been installed on more than 200 million Windows devices.

## #83 saturnian — Sep 8th, 2021, 09:46 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf
 Is it possible to add the management of the new DownloadStarting event of webview2 in RC6 ?.
 This would allow to change the default download directory.

 Thanks a lot

 François

 Here is the link to Microsoft Documentation of this new event
 <https://docs.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/icorewebview2_4?view=webview2-1.0.902.49> WebView2 Win32 C++ ICoreWebView2_4 | Microsoft Docs

## #84 softv — Sep 11th, 2021, 12:08 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Have just finished a Binding to the new WebView2-BrowserControl (based on Edge-Chromium).
 ... .. .
 ... .. .
 The new BaseDll-package of the RC6 now includes the official WebView2Loader.dll (version 1.0.674),
 which the cWebView2-class then works against.
 ... .. .
 ... .. .
 You should now be able to test this new Edge-Browser-Binding (even on Win7-OSes) via this little VB6-Demo:
 ... .. .
 Please let me know, when something is not working as expected -
 or when you want me to include a certain extra-functionality into the new cWebView2-class.

 I want to "finalize" the new RC6-functionality at the end of the year (then switching Binary-Compatibility on).

 Happy testing... *[img: Smilie]*

 Olaf

 Dear Olaf,

 First thing first. **Thanks a TONNN**, as ever. It was great seeing your demo in action! Wow!

 Actually, fortunate that I stumbled on this forum page (when I was really desperate to find a demo on how to use WebView2 in your RC6). How I wish all these demos of yours (or) links to these kinds of forum pages were in your vbRichClient.com website in one page! Anyway, I fully understand your lack of time.

 Well, you have asked us to let you know if something is not working. I am just giving hereunder what happened when I ran the demo:

 1. For the following procedures, VB6 reported "Compile error: Procedure declaration does not match description of event or procedure having the same name"

Code:

```
Private Sub WV_AcceleratorKeyPressed(ByVal KeyState As eWebView2AccKeyState, ByVal IsExtendedKey As Long, ByVal WasKeyDown As Long, ByVal IsKeyReleased As Long, ByVal IsMenuKeyDown As Long, ByVal RepeatCount As Long, ByVal ScanCode As Long, IsHandled As Long)
  Debug.Print "WV_AcceleratorKeyPressed"
End Sub
```

Code:

```
Private Sub WV_NavigationCompleted(ByVal IsSuccess As Long, ByVal WebErrorStatus As Long)
  Debug.Print "WV_NavigationCompleted"
End Sub
```

2. Instead of immediately researching on what the abovementioned error could be due to, I just commented out both the above procedures and ran the demo. And, there it was! Your demo in super action! **Thanks, Olaf!** People like you give lots of joy to persons like me!

 Well, why did the compile error pop up for the abovementioned two procedures? Have they to be redeclared with different number/types of parameters? If so, kindly let me know how. In case the Compile error was due to some other reason, then **kindly educate me** on the same too.

 Kind Regards.

## #85 jpbro — Sep 11th, 2021, 12:19 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **softv** *[img: View Post]*

Well, why did the compile error pop up for the abovementioned two procedures? Have they to be redeclared with different number/types of parameters? If so, kindly let me know how. In case the Compile error was due to some other reason, then **kindly educate me** on the same too.

See <https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)&p=5530453&viewfull=1#post5530453> https://www.vbforums.com/showthread....=1#post5530453. Short-answer RC6 is still in a development phase and the method signatures are not yet finalized.

## #86 softv — Sep 11th, 2021, 12:26 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

 ... .. .
 As for detection, whether a given target-system already contains this runtime, ... .. .
 ... .. .
 If you ask your user beforehand like this:

Code:

```
  If Len(WV.GetMostRecentInstallPath) = 0 Then 'Edge-Chromium-Webview is missing on this system
     If MsgBox("The needed MS-WebView2-runtime is missing," & vbLf & "Do you want to download the installer now?", vbYesNo) = vbYes Then
        New_c.FSO.ShellExecute "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
     End If
  End If
```

The little 1.7MB MS-Executable from that download above, will then determine which system the User actually runs (Win7, Win8, Win10) -
 and then download and install the larger (about 50MB-80MB) and properly matching evergreen-runtime which will work best on the given system.

 HTH

 Olaf

Thanks a LOT, Olaf! I was thinking of doing the detection myself via a tiny function but then I also thought some function must be already existing. And, there it is! Thanks.

 Kind regards.

## #87 softv — Sep 11th, 2021, 12:44 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thanks jpbro. Both the procedures work now!

 Kind regards.

## #88 saturnian — Nov 15th, 2021, 02:56 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello,

 Using RC6 cWebView2 with the latest Evergreen version of Webview2, on some sites (especially Google sites: Google News, Google Play Store ...) the DocumentURL and DocumentTitle properties remain an empty string once the page is completely displayed.
 This does not happen with old fixed versions of webview2 !
 Have you encountered this problem?
 Is this a Microsoft regression or an RC6-Webview interface problem or Google sites issues ?
 Best regards

 Note: I am using RC6 6.0.9 on Windows 10

## #89 Schmidt — Nov 16th, 2021, 09:16 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

Hello,

 Using RC6 cWebView2 with the latest Evergreen version of Webview2, on some sites (especially Google sites: Google News, Google Play Store ...) the DocumentURL and DocumentTitle properties remain an empty string once the page is completely displayed.
 This does not happen with old fixed versions of webview2 !
 Have you encountered this problem?
 Is this a Microsoft regression or an RC6-Webview interface problem or Google sites issues ?

I was able to reproduce that, and the reason is, that Google since a few weeks apparently insists (on some of their Sites) on "higher security" for the eval-function (when called from "self-assigned scripts" - which I use so far, to resolve window-Properties via jsProp Get/Let - as e.g. jsProp("document.title").

 Came up with a workaround for the Eval-Function (which formerly worked behind WV.jsProp Get/Let) -
 and it will be present in the next Release 6.0.10... (want to wait until SQLite 3.37 gets released at december 1'th)

 Olaf

## #90 saturnian — Nov 16th, 2021, 03:05 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

I was able to reproduce that, and the reason is, that Google since a few weeks apparently insists (on some of their Sites) on "higher security" for the eval-function (when called from "self-assigned scripts" - which I use so far, to resolve window-Properties via jsProp Get/Let - as e.g. jsProp("document.title").

 Came up with a workaround for the Eval-Function (which formerly worked behind WV.jsProp Get/Let) -
 and it will be present in the next Release 6.0.10... (want to wait until SQLite 3.37 gets released at december 1'th)

 Olaf

thank you for these details, Olaf
 I will wait for RC6 version 6.0.10.
 Thank you again for this fantastic work !

 François

## #91 I. Capri — Nov 19th, 2021, 07:52 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

// Deutsch
 Sehr geehrter Herr Schmidt,

 ich hätte gern gewusst, ob Sie das webResourceRequested-Event in RC6.DLL zur Verfügung stellen können, sowie wie Sie ans ICoreWebView2-Interface dran kommen können.

 Vielen Dank im Voraus für Ihre Rückmeldung!

 Mit freundlichen Grüßen
 Igli Kapri

 // Englisch
 Dear Mr. Schmidt,

 I would like to know whether You can provide the webResourceRequested-Event in RC6.DLL as well as how You find the ICoreWebView2 interface in .NET.

 Thank you very much in advance for Your response!

 Kind regards,
 Igli Kapri

## #92 workmail — Nov 21st, 2021, 01:26 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

vb6+cwebview2 is good thing

 but did somebody know why could not upload file in vb6 cwebview2 edge browser?

 some web could, but most important web could not use web file upload system. What a pity.*[img: Name:  fcff.jpg
Views: 6465
Size:  12.9 KB]*

## #93 talatoncu — Nov 27th, 2021, 02:34 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

How can I click a button on the page?

 I have tried

 WV.jsRun ("document.getElementById('btnSubmit').click")

 But nothing happened

## #94 WZabel — Nov 27th, 2021, 02:56 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

In WebViewInit
 'we can predefine our own set of js-functions, before any document gets loaded
 WV.AddScriptToExecuteOnDocumentCreated "function ElementClick(myElement){ window.document.getElementById(myElement).click(); }"

 'after document load we can use
 WV.jsRun "ElementClick", "btnSubmit"

## #95 Schmidt — Nov 27th, 2021, 03:29 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **talatoncu** *[img: View Post]*

 WV.jsRun ("document.getElementById('btnSubmit').click")

jsRun is thought for function-calls, which allow you to pass parameters as normal VB-Vars.

 For such longer "multi-dotted-expressions" (with "inner, literal-arguments"), you are better advised to use:
 WV.ExecuteScript "document.getElementById('btnSubmit').click()"

 Please note the trailing parentheses behind the last methodname "click".

 Olaf

## #96 talatoncu — Nov 28th, 2021, 01:13 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

jsRun is thought for function-calls, which allow you to pass parameters as normal VB-Vars.

 For such longer "multi-dotted-expressions" (with "inner, literal-arguments"), you are better advised to use:
 WV.ExecuteScript "document.getElementById('btnSubmit').click()"

 Please note the trailing parentheses behind the last methodname "click".

 Olaf

Perfect!

 Thank you very much Sir!

 You saved my life with this implementation and your immediate answer.

## #97 EasyOneX — Dec 28th, 2021, 01:08 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi.

 Unfortunately, due to the security feature of the Edge Browser, I cannot access local files, e.g. something like this: file: /// C: /Pictures/loading.gif

 Error message in the console: Not allowed to load local resource

 How can I change the security settings to allow local resources?

 thx Andreas

## #98 Schmidt — Dec 28th, 2021, 05:53 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **EasyOneX** *[img: View Post]*

Unfortunately, due to the security feature of the Edge Browser, I cannot access local files, e.g. something like this: file: /// C: /Pictures/loading.gif

 Error message in the console: Not allowed to load local resource

 How can I change the security settings to allow local resources?

There's a huge deal of influence possible, via the CommandLine-Parameter of the intial Binding-Call...

 I've posted a Link to it already in an earlier posting here in this thread,
 but here is it again (already jumping to a setting which might work for you):
 <https://peter.sh/experiments/chromium-command-line-switches/#allow-file-access-from-files> https://peter.sh/experiments/chromiu...ess-from-files

 If this specific one doesn't, other flags might do the job (perhaps google a bit about it,
 using the term CEF in your searches).

 Olaf

## #99 EasyOneX — Dec 29th, 2021, 05:51 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

There's a huge deal of influence possible, via the CommandLine-Parameter of the intial Binding-Call...

 I've posted a Link to it already in an earlier posting here in this thread,
 but here is it again (already jumping to a setting which might work for you):
 <https://peter.sh/experiments/chromium-command-line-switches/#allow-file-access-from-files> https://peter.sh/experiments/chromiu...ess-from-files

 If this specific one doesn't, other flags might do the job (perhaps google a bit about it,
 using the term CEF in your searches).

 Olaf

Perfect. I hadn't seen the page with the commandlines yet. That's great. Thank you very much and happy holidays I wish you...

 Thx Andras

## #100 EasyOneX — Dec 29th, 2021, 06:18 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thanks for the link. I hadn't seen the website yet. I wish many dear and happy holidays...

## #101 xxdoc123 — Dec 30th, 2021, 12:16 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I know that using chromedirver. can intercept the image of the specified element
 By the following command
 /session/" & m_StrSessionId & "/element/" & m_StrElementId & "/screenshot.

>

>
 Private Sub cmdCaptureWV_Click()
 Dim Srf As cCairoSurface
 Set Srf = WV.CapturePreview(CaptureAs_PNG) 'capture the current WV-Window as a Cairo-Image-Surface
 Srf.WriteContentToPngFile App.Path & "\WV_Capture.png" 'which we can now visualize, or just write out as a PNG-file
 End Sub

Do you support this function?

## #102 Schmidt — Dec 30th, 2021, 04:12 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xxdoc123** *[img: View Post]*

I know that using chromedirver. can intercept the image of the specified element
 By the following command
 /session/" & m_StrSessionId & "/element/" & m_StrElementId & "/screenshot.

 Do you support this function?

No, there's no Element-based ScreenShot-feature (what's built-in, is working Window-based).
 One could get around making that "bigger ScreenShot" -
 then getting the surrounding rect of the Element via javascript -
 followed by cutting out the Sub-Bitmap using these coords...

 But if this is for Unit-Testing... that's what chromedriver is made for and specialized in
 (not only with regards to the ScreenShot-feature) ... so, why not use it furhter?

 Olaf

## #103 xxdoc123 — Dec 30th, 2021, 11:02 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

We know that Google Chrome can set the capabilities parameter to customize the functions of the browser through chromdirver. I don’t know if you have this setting.For example: disable camera ...

 I want to replace the js function loaded in the original webpage.how can i do
 Originalfunnctio is checkCamera

 WV.ExecuteScript "checkCamera= function myfunction(){}" ?

## #104 xxdoc123 — Dec 30th, 2021, 11:11 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

We know that Edge-Chromium can set the capabilities parameter to customize the functions of the browser through Edgedirver. I don’t know if you have this setting.For example: disable camera ...

 I want to replace the js function loaded in the original webpage.how can i do
 Originalfunnctio is checkCamera

 WV.ExecuteScript "checkCamera= function myfunction(){}" ?

## #105 Schmidt — Dec 31st, 2021, 02:26 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xxdoc123** *[img: View Post]*

I don’t know if you have this setting.For example: disable camera ...

As already said a few posting further above - there is a ton of startup-settings,
 you can set and combine via the commandline-parameter of the initial Bind-Call.

 Here is the link again: <https://peter.sh/experiments/chromium-command-line-switches/> https://peter.sh/experiments/chromiu...line-switches/
 (if you search via <Ctrl><F>, you will find a lot things to influence "media" or "capture"-behaviour.)

 Olaf

## #106 xxdoc123 — Dec 31st, 2021, 03:30 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

thanks.

 I am not very proficient in JS.I have to try to use Js.

## #107 meco_r — Jan 5th, 2022, 06:28 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf and thanks for your great work!
 I have a "stupid issue" to report regarding UserContextMenu event: the ScreenY argument declaration is spelled wrong: "SreenY" *[img: Smilie]*
 Someting more serious: are you planning to map the NavigationStarting Event in the vb6 wrapper so that is possible to detect the destination url before the navigation effectively occurs?

 Ciao,
 DR

## #108 Resurrected — Jan 12th, 2022, 03:33 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf

 I have a question regarding RC6.WebView2.

 Is there a way to download an image resource instead of opening it in WebView2? By using Navigate, the image is opened in WebView2 instead of invoking the download operation. I'd like that image to be downloaded the way I navigate to a zip file.

 Well, I know I could get the image url and use a separate httpclient to download the file. But anyway to use the WebView2?

 Plus it doesn't seem the RC6.WebView2 provide anything related to downloading.

 Thanks in advance.

## #109 C--E — Jan 18th, 2022, 11:02 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello,

 I downladed RC6 a few days ago from vbRichClient.com and I registered RC6.dll with the aid of RegisterRC6inPlace.vbs.
 Then I installed the "evergreen-WebView2-runtime" via the MS-download-link.

 I tried to run WebView2Demo.vbp, got the error "Procedure declaration does not match description ..." and corrected the code as explained in post #60.

 Now I expected the demo to run, but it always stops in the following line:

Code:

```
If WV.BindTo(picWV.hWnd) = 0 Then MsgBox "couldn't initialize WebView-Binding": Exit Sub
```

with the following error message:
 **"Runtime Error 53:
 Error in FindFirstFile: ErrNum: 3, The system cannot find the path specified."**

 (I'm using VB6 Version 8176 under Win10Pro).

 Any ideas what is wrong or what I did wrong?

## #110 Schmidt — Jan 18th, 2022, 01:05 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **C--E** *[img: View Post]*

Hello,

 I downladed RC6 a few days ago from vbRichClient.com and I registered RC6.dll with the aid of RegisterRC6inPlace.vbs.
 Then I installed the "evergreen-WebView2-runtime" via the MS-download-link.

 I tried to run WebView2Demo.vbp, got the error "Procedure declaration does not match description ..." and corrected the code as explained in post #60.

 Now I expected the demo to run, but it always stops in the following line:

Code:

```
If WV.BindTo(picWV.hWnd) = 0 Then MsgBox "couldn't initialize WebView-Binding": Exit Sub
```

with the following error message:
 **"Runtime Error 53:
 Error in FindFirstFile: ErrNum: 3, The system cannot find the path specified."**

 (I'm using VB6 Version 8176 under Win10Pro).

 Any ideas what is wrong or what I did wrong?

There seems to be a problem with the installation of the WebView-Runtime...

 The BindTo-call is looking for the WebView-Runtime-BaseFolder via the following Path-Expression:
 (for your Immediate-Window):
 ?New_c.FSO.GetSpecialFolder(CSIDL_PROGRAM_FILESX86) & "\Microsoft\EdgeWebView\Application"

 When I paste the above into my immediate-window (followed by pressing <Return>),
 then it prints out:
 C:\Program Files (x86)\Microsoft\EdgeWebView\Application

 This path exists on my machine (when I paste it into Explorer.exe, it shows the content)...

 Under this path, the BindTo-call then continues, to list determine the Folder with the highest Version-Number...

 So, could you check what your immediate-window says, the Base-InstallPath of the WebView-runtime is -
 and whether it exists on your machine?

 HTH

 Olaf

## #111 C--E — Jan 18th, 2022, 05:29 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The output of the Immediate Window is: C:\Program Files (x86)\Microsoft\EdgeWebView\Application
 This path does not exist on my machine!

 Some minutes ago I re-installed "evergreen-WebView2-runtime" via the MS-DownloadLink, which took place without problems (which was also the case during the first installation), and now the path \EdgeWebView\Application exists and the demo program is running!

 Thank you very much for your quick help!

## #112 xxdoc123 — Jan 21st, 2022, 06:52 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Save As is not available in web pages

## #113 Schmidt — Jan 21st, 2022, 07:53 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xxdoc123** *[img: View Post]*

Save As is not available in web pages

"Save As..." is available in the Default-Context-menu...
 and it will work, when you do not run elevated as an Administrator (as is the case e.g. in the IDE)...

 When you run a compiled version, it should pop-up the FileDialog (for you to choose a Save-target).

 (MS knows about the problem in Admin-Mode, there is a ticket for it: <https://github.com/MicrosoftEdge/WebView2Feedback/issues/802> https://github.com/MicrosoftEdge/Web...ack/issues/802)

 HTH

 Olaf

## #114 xxdoc123 — Jan 22nd, 2022, 06:07 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

thanks

## #115 C--E — Jan 23rd, 2022, 12:56 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Once again thanks to Olaf for his help and generally for his great work!

 Now I'm succesful with the WebView2Demo.vbp, added a second picturebox named "Pchart" and wanted to copy the let's say "browsing result", i.e. the image in picWV, into this second picturebox. I tried:
 PChart.Picture = picWV.Image
 but obviously that doesn't work.
 Any ideas on how to get this done?

## #116 Schmidt — Jan 23rd, 2022, 02:31 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **C--E** *[img: View Post]*

PChart.Picture = picWV.Image
 but obviously that doesn't work.

The PictureBox picWV which is used in the Demo, is only providing "a Container-hWnd" for the Browser-instance,
 (which places its own Browser-hWnd as a Child inside the picWV.hWnd).

 But there is a builtin Method, which captures the current Browser-Content as a Pixel-Graphic
 (returning a cCairoSurface-Object, which you can use for all kind of purposes, e.g. writing to a PNG-ByteArray or -File for example)

 Long story short - please try:
 Set PChart.Picture = WV.CapturePreview(CaptureAs_PNG).Picture

 Olaf

## #117 C--E — Jan 23rd, 2022, 05:20 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I saw the following part in the demo:
  Dim Srf As cCairoSurface
 Set Srf = WV.CapturePreview(CaptureAs_PNG) 'capture the current WV-Window as a Cairo-Image-Surface
 Srf.WriteContentToPngFile App.Path & "\WV_Capture.png" 'which we can now visualize, or just write out as a PNG-file
 but was not able to transform this into the syntax:
 Set PChart.Picture = WV.CapturePreview(CaptureAs_PNG).Picture

 Long story short: You code works like a charm - thank you very much!!!

 Best Regards
 C-E

## #118 Schmidt — Jan 24th, 2022, 12:57 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **C--E** *[img: View Post]*

I saw the following part in the demo:
 ...was not able to transform this into the syntax:
 Set PChart.Picture = WV.CapturePreview(CaptureAs_PNG).Picture

The above "one-liner" is just making use of a compiler-feature,
 which creates a (hidden) temporary cCairoSurface-instance under the covers,
 (from the return-value of the CapturePreview-method) before accessing its Picture-Property...

 The PNG-output code could have been written in one line as well -
 I just wanted to make it more obvious, what CapturePreview returns.

 A cCairoSurface is just a wrapper-class around a "32Bit-Per-Pixel DIB-allocation" in system- (not GPU-) memory -
 but also the entry-point into powerful cairo-based (Vector-)Drawing-Commands, which are just "one context-call" away.

 E.g. you could easily draw "additional stuff on it" (before placing such a Surface on a PictureBox),
 for example - if you want to "frame" the Browser-ScreenShot, you could do it this way:

Code:

```
  Dim Srf As cCairoSurface 'make the returned Surface explicit
  Set Srf = WV.CapturePreview(CaptureAs_PNG)
  With Srf.CreateContext 'enter Drawing-Mode
    .SetLineWdth 2
    .Rectangle 0, 0, Srf.Width, Srf.Height, True 'define a Rectangle-Path
    .SetSourceColor vbRed, 0.5 'define the Border-Color with 50% Alpha
    .Stroke 'strokes only the Outline of the Path (not the interiour of the Rect)
  End With
  Set PChart.Picture = Srf.Picture 'now assign the combined result to the PicBox
```

HTH

 Olaf

## #119 EasyOneX — Feb 1st, 2022, 08:27 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi, everyone.

 Also, is it possible to auto-fit a text box to the available width and height so that when the shape is resized, the text box will also resize proportionally to the available area?

 From this :

Code:

```
<textarea id=txt1 cols=50 rows=4></textarea>
```

to this:

Code:

```
<textarea id=txt1 height:auto width:auto></textarea>
```

Kind regards

 Andrew

## #120 Bearfot — Feb 3rd, 2022, 02:42 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

@Olaf First of all fantastic work!! One question:
 Would it be possible to extend cWebView2 with a callback when basic authentication is requested from a site so I can provide with user name and password and thereby prevent the Sign in prompt? I.e. my hybrid app first requires a login by the user, and when navigating to my site an additional "Sign in" is currently now needed. Earlier when using the embedded IE control this was automatically handled by WinInet.

 Thanks in advance

## #121 Bearfot — Feb 4th, 2022, 02:45 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi again,
 what is the difference between providing arguments in the New_c.WebView2(...) compared to BindTo(...)?
 I.e. adding userDataFolder and additionalBrowserArguments seems only have to affect in BindTo(), and also it seems that even the arguments are optional I can't just provide the additionalBrowserArguments without also specifying a userDataFolder. If userDataFolder is left empty or as an empty string, BindTo will fail.

 regs
 Stefan

## #122 Schmidt — Feb 5th, 2022, 03:53 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Bearfot** *[img: View Post]*

what is the difference between providing arguments in the New_c.WebView2(...) compared to BindTo(...)?

None ...
 The only difference is, that the optional timeout-param is a bit shorter (5sec) behind New_c, than in BindTo (8sec) -
 though this might exactly be the difference on some older (slower) machines between the two calls.
 I have corrected that timeout-difference in the cConstructor-call to 8sec now (to match with BindTo).

>

>

 *[img: Quote]* Originally Posted by **Bearfot** *[img: View Post]*

I.e. adding userDataFolder and additionalBrowserArguments seems only have to affect in BindTo(),

All the Params are just "routed through" to BindTo, as they come in via New_c.WebView2 -
 (the only exception - as said - being a different timeout, when you leave this param-slot free).

>

>

 *[img: Quote]* Originally Posted by **Bearfot** *[img: View Post]*

and also it seems ...
 I can't just provide the additionalBrowserArguments without also specifying a userDataFolder.
 If userDataFolder is left empty or as an empty string, BindTo will fail.

If userData is left empty, the BindTo-call will internally not "leave it that way", but initialize it to:
 If Len(userDataFolder) = 0 Then userDataFolder = New_c.FSO.GetLocalAppDataPath
 (which resolves e.g. on my Win10 to: C:\Users\Olaf\AppData\Local)

 Perhaps certain commandline-arguments only work with certain UserDataFolder-Paths
 (which might also depend on the OS-version you are using - I've heard of Users who have had problems on Win7)...

 As for your earlier question... I think "finding the right HTML5" (and Css- or style-attributes)
 is in the responsibility of the Users of this control-binding (to not blow this thread here "out of proportion") -
 and would suggest to post your HTML-questions in a better matching Forum (or SubForum on this site).

 HTH

 Olaf

## #123 Bearfot — Feb 5th, 2022, 05:10 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

@Olaf, thanks for reply!

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

As for your earlier question... I think "finding the right HTML5" (and Css- or style-attributes)
 is in the responsibility of the Users of this control-binding (to not blow this thread here "out of proportion") -
 and would suggest to post your HTML-questions in a better matching Forum (or SubForum on this site).

My earlier question was regarding basic authentication. If it was possible to get a callback from cWebView2 and provide user name and password in order to prevent the Sign in prompt (if a site uses basic authentication). E.g. something like GetCredentials(ByVal Url As String, UserName As String, Password As String, Cancel As Boolean).

 regs
 Stefan

## #124 Schmidt — Feb 5th, 2022, 06:08 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Bearfot** *[img: View Post]*

My earlier question was regarding basic authentication.

Oops, my bad (interchanged your post with that from Andrew/EasyOneX)...

>

>

 *[img: Quote]* Originally Posted by **Bearfot** *[img: View Post]*

If it was possible to get a callback from cWebView2 and provide user name and password in order to prevent the Sign in prompt (if a site uses basic authentication). E.g. something like GetCredentials(ByVal Url As String, UserName As String, Password As String, Cancel As Boolean).

I don't know what exact behaviour you relate to, when you talk about "WinInet-usage" in conjunction with the old IE-Control.

 If you "already know" the user-credentials from an earlier "logon-attempt"...
 then the basic-auth credentials can be passed directly in the URL, as e.g.:

Code:

```
<http://username:password@example.com/> "http://username:[email protected]/"
```

There's also a lot of different "auth"-related params for the commandline... you might want to experiment with.

 And finally, there's (since 6.0.9) a new method:
 - NavigateWithWebResourceRequest (which has an optional param: Headers_CrLfSeparated As String)
 which will allow you to specify headers like e.g.: "Authorization: Basic <your base64-encoded credentials>"

 HTH

 Olaf

## #125 Bearfot — Feb 5th, 2022, 07:57 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Yes, passing username*[img: Stick Out Tongue]*assword in the url is how I now currently solved it, but that is not a preferable solution from a security perspective.
 Instead username/password should not be provided unless the site requests for it, and with a callback I can check if I have the proper credentials for the url, and in that case provide them or otherwise set cancel = true which will bring up the default Sign in prompt. Just a proposal for a nice feature for the future :-)

 Thanks again
 Stefan

## #126 xxdoc123 — Feb 8th, 2022, 08:49 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

When I execute the js command,
 WV.AddScriptToExecuteOnDocumentCreated jscode
 WV.jsRunAsync "GetLinks", "div.picList > ul > li>a" '. error Unable to read in WV_JSAsyncResult. .

 if changed to

 WV.AddScriptToExecuteOnDocumentCreated jscode
 sleep 1000
 WV.jsRunAsync "GetLinks", "div.picList > ul > li>a" '.
 execution succeed. Is there a better way without sleep?

 Also execute WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();" sometimes the webpage is not executed successfully. So I will modify it to
 WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();"
 sleep 1000
 WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();"
 sleep 1000
 WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();"

 Is there any good way to wait for js to finish executing synchronously?

 How is the jsCallTimeOutSeconds function used?

## #127 avinash7 — Feb 10th, 2022, 09:52 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,

 Thank you for such a great contribution. I was able to achieve most of the task that I need to.

 However, I am unable to handle below scenarios:

 1. Handle multiple downloads: I want to enable this by default to allow multiple file download (disable multiple file download confirmation). This is because I perform very lengthy process which ends with downloading multiple files which is currently blocked after downloading first file.

 2. Default file download location: How to specify default file download location?

 I tried to play with %appdata%\EBWebView\Default\Preference file that stores this settings. However, this file is overwritten automatically by some process every few seconds which causes reversal of my changes.

 Also, I see that WebView2 has **DownloadStarting** event that allows to achieve some part of above, but I do not see that available in cWebView2.

## #128 Bearfot — Feb 13th, 2022, 12:30 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,
 if you have any plans of extending the cWebView2 events (e.g. DownloadStarting) then also the **DocumentTitleChange** is of interest. I have solved it now by adding the following code in the DocumentComplete event:
 m_webView.ExecuteScript "document.addEventListener('DOMSubtreeModified', function(e){if(e.target.nodeName === 'TITLE') vbH().RaiseMessageEvent('OnTitleChange',e.target.text)})"
 This is maybe not the most efficient solution, but at least it works.

 In addition I'm also very interested in getting access to "CoreWebView2EnvironmentOptions". Its property "AdditionalBrowserArguments" can be applied in the constructor and BindTo method, but is is possible to access the "Language" and "AllowSingleSignOnUsingOSPrimaryAccount" properties?

 Finally in order to solve my earlier WinInet issues there seems to be a solution by copying cookies from WinInet to WebView2, but it would the require access to the "CoreWebView2CookieManager" object. Any plans to make such extensions of the cWebView?

 regs
 Stefan

## #129 Bearfot — Feb 13th, 2022, 01:29 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi again Olaf,
 one more question. I found a bug when adding an object using the AddObject() method which I guess internally calls the "AddHostObjectToScript". It seems that you can't access an attribute named "length" of such object. The solution I found was to add the following line for the created javascript proxy object:
 var myobject = window.chrome.webview.hostObjects.sync.myobject; // myobject with a "length" attribute added with cWebView2.AddObject()
 alert(myobject.length); // will always display 0
 // adding the following solves the issue
 myobject._forceRemoteProperties.add("length"); // myobject is actually a proxy object in javascript
 alert(myobject.length); // now shows the correct value of myobject.length

 As I don not know the implementation details, is this a bug in cWebView2.AddObject() or something I should report in github for CoreWebView2.AddHostObjectToScript()?

 regs
 Stefan

## #130 saturnian — Feb 18th, 2022, 06:06 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello avinash7 and Bearfot

 Olaf seems to rarely respond to requests to add functionality to RC6. It's probably to avoid putting pressure on himself. The task he has set himself is indeed hudge.
 In a recent conversation, he told me he hoped to integrate Webview2 download management in version 6.0.10 of RC6. I look forward to this version to continue my developments. I even check every day if the new version is posted on his site!
 We'll have to be patient *[img: Smilie]*
 Thanks again to Olaf for his fantastic work !

## #131 Bearfot — Feb 18th, 2022, 06:54 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The AddHostObjectToScript issue is confirmed bug in webview2 (issue #2185) so this can be ignored.
 Access of the cookie manager can also be skipped as this seems not solve my issue.
 Other things just nice to have, so no remaining issues from me.

## #132 avinash7 — Feb 19th, 2022, 03:34 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

 It's probably to avoid putting pressure on himself

Sorry Olaf and saturnian if it looked like creating pressure. I posted here to see if there is anything already available and can be leveraged to achieve what I wanted to do. This is because I see that Olaf has surfed breadth and depth of this area!

## #133 BhargaviU — Mar 4th, 2022, 08:19 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Have just finished a Binding to the new WebView2-BrowserControl (based on Edge-Chromium).

 I've included this Binding (all in a single Class, named cWebView2) in the new RC6-version of the RichClient-lib
 (please download this new version 6 from its usual place, at vbRichClient.com first).

 The new BaseDll-package of the RC6 now includes the official WebView2Loader.dll (version 1.0.674),
 which the cWebView2-class then works against.

 Please note, that the above Binding will currently require, that you install the larger:
 "Evergreeen WebView2-Runtime" (not included in the RC6-BasePackage).
 Here the official MS-DownloadLink for the evergreen-installer: <https://go.microsoft.com/fwlink/p/?LinkId=2124703> https://go.microsoft.com/fwlink/p/?LinkId=2124703

 So, after ensuring the mentioned two prerequisites:
 - the Dlls of the new RC6-package in a folder of your choice + a registered RC6.dll
 - and the successfull installation of the "evergreen-WebView2-runtime" via the MS-download-link above

 You should now be able to test this new Edge-Browser-Binding (even on Win7-OSes) via this little VB6-Demo:
 <https://www.vbforums.com/attachment.php?s=9b84d96bf559669a1dc45b4f80d160b8&attachmentid=179166> Attachment 179166

 Please let me know, when something is not working as expected -
 or when you want me to include a certain extra-functionality into the new cWebView2-class.

 I want to "finalize" the new RC6-functionality at the end of the year (then switching Binary-Compatibility on).

 Happy testing... *[img: Smilie]*

 Olaf

Hi Schmidt,
 I tired WebView2Demo it is really great. I wanted to know if we can integrate events of webbrowser1 with its replacements in WebView2.
 For Example :
 1.StatusTextChange
 2.BeforeNavigate2
 3.WindowClosing

 I tried few events but was unable to recreate the events I mentioned before.
 Could you please guide me for the same.

## #134 LeandroA — Mar 6th, 2022, 03:08 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi is it possible to get an object with jsProp, getElementsByClassName doesn't work for me either.

Code:

```
Dim Elements As Object
Set Elements = Document.GetElementsbyClassName("FTBzM")
```

with IE this worked, should I think about it differently?

## #135 BooksRUs — Mar 10th, 2022, 06:29 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I'm new to all this, so take it easy on me, please *[img: Smilie]*

 I have successfully used the WebView2 in the VB6 IDE without issues... BUT, when I go to MAKE my project, it fails to compile with the following error:

 **Unexpected error occurred in code generator or linker. --View error messages?**

 I click on Yes and get a text tmp file that contains this error:

 C:\<MYFOLDER>\<MYFORM>.frm(6624) : **fatal error C1001: INTERNAL COMPILER ERROR**
 ** (compiler file 'e:\work\utc2\src\P2\main.c', line 495)**
 ** Please choose the Technical Support command on the Visual C++**
 ** Help menu, or open the Technical Support help file for more information**

 If I comment out any lines where I'm doing this:

 **Set wv5 = New_c.WebView2 **

 or this:

 **If wv5.BindTo(... blah, blah... ) then**

 Then the program compiles fine. Obviously, it doesn't *work* correctly because these lines are commented out.

 Again, if I just Run the program in the VB6 IDE, then it works great.

 Please help.

## #136 SearchingDataOnly — Mar 11th, 2022, 08:11 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

I'm new to all this, so take it easy on me, please *[img: Smilie]*

 I have successfully used the WebView2 in the VB6 IDE without issues... BUT, when I go to MAKE my project, it fails to compile with the following error:

 **Unexpected error occurred in code generator or linker. --View error messages?**

 I click on Yes and get a text tmp file that contains this error:

 C:\<MYFOLDER>\<MYFORM>.frm(6624) : **fatal error C1001: INTERNAL COMPILER ERROR**
 ** (compiler file 'e:\work\utc2\src\P2\main.c', line 495)**
 ** Please choose the Technical Support command on the Visual C++**
 ** Help menu, or open the Technical Support help file for more information**

 If I comment out any lines where I'm doing this:

 **Set wv5 = New_c.WebView2 **

 or this:

 **If wv5.BindTo(... blah, blah... ) then**

 Then the program compiles fine. Obviously, it doesn't *work* correctly because these lines are commented out.

 Again, if I just Run the program in the VB6 IDE, then it works great.

 Please help.

Maybe it's a question about the RC6 path.
 <https://www.vbforums.com/showthread.php?882171-Multi-Process-RPC-Listener-Host-System-for-RC5&p=5443031&viewfull=1#post5443031> https://www.vbforums.com/showthread....=1#post5443031

## #137 BooksRUs — Mar 11th, 2022, 11:17 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

Maybe it's a question about the RC6 path.
 <https://www.vbforums.com/showthread.php?882171-Multi-Process-RPC-Listener-Host-System-for-RC5&p=5443031&viewfull=1#post5443031> https://www.vbforums.com/showthread....=1#post5443031

 I used the included RegisterRC6inPlace.vbs to register RC6 on my dev PC, directly inside my App folder, in hopes that nothing would be confused as to where anything was.

 I did extract all the DLL files into this same folder. Does any of the other DLLs need to be registered just so the VB6 can compile an EXE that uses RC6?

 Thanks!

## #138 Schmidt — Mar 12th, 2022, 08:13 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

I used the included RegisterRC6inPlace.vbs to register RC6 on my dev PC, directly inside my App folder, in hopes that nothing would be confused as to where anything was.

 I did extract all the DLL files into this same folder. Does any of the other DLLs need to be registered just so the VB6 can compile an EXE that uses RC6?

 Thanks!

If it runs in the IDE, then there is nothing wrong with the registration of RC6 -
 so the compiler-error you see has to have other reasons - like e.g.:
 - you're not running the IDE "as Admin" (when compiling)
 - you're not using the most recent ServicePack for the VB6-IDE
 - you've registered the RC6 on a Network-Path, or a "mapped Network-Drive"
 - you've opened your VB-Project from a Network-Path or a "mapped Network-Drive"
 - you're trying to compile the Binary "into a Network-Path or a mapped Network-Drive"

 I'd start trying to fix this, by ensuring Admin-Mode for the IDE -
 and try to "open, and compile to" from/to local Folders on local Disks...
 (also making sure, that "compiling Debug-info into your Executable" is deactivated in the IDEs compiler-settings)

 HTH

 Olaf

## #139 BooksRUs — Mar 12th, 2022, 11:56 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

If it runs in the IDE, then there is nothing wrong with the registration of RC6 -
 so the compiler-error you see has to have other reasons - like e.g.:
 - you're not running the IDE "as Admin" (when compiling)
 - you're not using the most recent ServicePack for the VB6-IDE
 - you've registered the RC6 on a Network-Path, or a "mapped Network-Drive"
 - you've opened your VB-Project from a Network-Path or a "mapped Network-Drive"
 - you're trying to compile the Binary "into a Network-Path or a mapped Network-Drive"

 I'd start trying to fix this, by ensuring Admin-Mode for the IDE -
 and try to "open, and compile to" from/to local Folders on local Disks...
 (also making sure, that "compiling Debug-info into your Executable" is deactivated in the IDEs compiler-settings)

 HTH

 Olaf

 First, I want to say a big THANKS to Olaf, for all your hard work on all of these tools! Also, for your time to jump in and try to help me out. Right away, *none* of the suggestions you have apply, I double-checked them all.

 The particular form I'm adding this to is the "main" form for a pretty large program that has been my project for the last 13 years. It is on the edge of having more controls than VB6 allows, so I thought perhaps having to declare WV (with events) was pushing it over the limit. I don't know if it's possible, but it would be wonderful if we could have an ARRAY of WV (with events). I'm not sure if these declarations "count" towards the VB6 limits or not, but either way, this would allow an array of WV that match an array of PicBoxes that "host" the WVs that I want to use the in program (there are several). This is exactly how I'm implementing this with the Webbrowser controls that I'm trying to replace.

 Anyway, I went "back to basics" and tried to compile the sample program and it compiled right up without problems AND it actually works when Run, both in the IDE and the EXE.

 So, I *moved* these statements:
 Set wv5 = New_c.WebView2
 If wv5.BindTo(FZwv(5).hwnd, , "C:\<myfolder>\EdgeWebView") = 0 Then

 ... to the Form_Load event (as in the Demo program) and the program now *does* Compile, but it does *not* work. For some reason, in my program, it seems to want me to do the BindTo again in order to work correctly...

 Perhaps, this is because the PicBox isn't really visible YET in the Form_Load event? Does the PicBox actually *have* to be Visible in order for the BindTo to work correctly? The main form isn't yet visible in the load event, it's busy loading... AND the user has yet to actually do anything that will make the WV/PicBox visible in the program, as this is an option for the user to turn this on or off.

 Again, Thanks for all your work on this and other tools over the years.

## #140 wwolf — Mar 12th, 2022, 03:10 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>
 Perhaps, this is because the PicBox isn't really visible YET in the Form_Load event? Does the PicBox actually *have* to be Visible in order for the BindTo to work correctly? The main form isn't yet visible in the load event, it's busy loading... AND the user has yet to actually do anything that will make the WV/PicBox visible in the program, as this is an option for the user to turn this on or off.

That's right! The window to whose hwnd the WebView is bound must be visible. This can be quite tricky with modal windows.

 In my sample code, I do something like this: <https://www.vbforums.com/showthread.php?895441-Liquidity-planner-new-WebView2-demo-application> https://www.vbforums.com/showthread....mo-application

## #141 BooksRUs — Mar 12th, 2022, 09:51 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **wwolf** *[img: View Post]*

That's right! The window to whose hwnd the WebView is bound must be visible. This can be quite tricky with modal windows.

 In my sample code, I do something like this: <https://www.vbforums.com/showthread.php?895441-Liquidity-planner-new-WebView2-demo-application> https://www.vbforums.com/showthread....mo-application

It appears that the entire project is not based on normal VB6 forms. Without a total rewrite, are there some snippets of code that I could possibly use to get passed the Visible requirement, or to trick it into working?

 Thanks!

## #142 Schmidt — Mar 13th, 2022, 08:13 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

Perhaps, this is because the PicBox isn't really visible YET in the Form_Load event?

Yes, the PicBox.hWnd (any hWnd) you bind to,
 has to be visible at the point you call WV.BindTo...

 And any hWnd-based Child-Control will report False (regarding its visibility),
 as long as its hosting TopLevel-Parent is still invisible (in your case, your Form).

 Please take another look at the first line in my original Demo-Codes Form_Load (and the comments I made there).

 Olaf

## #143 BooksRUs — Mar 14th, 2022, 08:37 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Yes, the PicBox.hWnd (any hWnd) you bind to,
 has to be visible at the point you call WV.BindTo...

 And any hWnd-based Child-Control will report False (regarding its visibility),
 as long as its hosting TopLevel-Parent is still invisible (in your case, your Form).

 Please take another look at the first line in my original Demo-Codes Form_Load (and the comments I made there).

 Olaf

 Thanks again, Olaf. I spent alot of time trying different things, but I feel like I'm spinning my wheels because I was getting away from the *real* problem. I can live within the parameters of when to be Visible and when to SET & BINDTO, as the program is actually *working*.

 The real problem is that the program won't *compile*. So, I tried different settings on the VB6 compiler dialog -- it appears that the only way I can get it to compile is to change it to Compile to P-Code. Any other setting I tried with a working program, would generate the compiler error in the original post.

 Only time will tell if I'm able to keep these settings and distribute without causing other unknown issues, but I'm going to give it a try.

 Thanks again for helping me out. If there's anything else I can provide that might help solve the issue, just let me know.

 *[img: Smilie]*

## #144 BooksRUs — Mar 15th, 2022, 08:57 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

WV.jsRun is thought for simple (for the most part, self-defined) functions...

 For more complex "actions" like yours above, try it with:
 WV.ExecuteScript "document.getElementById('Somebutton').click()"
 instead.

 HTH

 Olaf

 I'm trying to simply automatically login to a website. I'm replacing a VB6 webbrowser control. I have already analyzed the login page and I know the fields I want to fill in are "user_name" and "user_password". The login button is on the first (0) form on the page. So, I used to do this:

 set w=webbrowser.Document
 set f=w.forms.item(0)
 set e=f.elements
 e.Item("user_name").Value = <myuserid>
 e.Item("user_password").Value = <mypassword>
 f.submit

 My question is, how do I do this with the new WebView since the DOM is not there, and the Submit button doesn't *have* a name?

 Thanks!

## #145 Schmidt — Mar 16th, 2022, 05:33 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

My question is, how do I do this with the new WebView since the DOM is not there...

The DOM is there, but you have to interact with it in js now...
 (cWebView2 is making this easy via AddScriptToExecuteOnDocumentCreated)

Code:

```
With New_c.StringBuilder

    .AddNL "function submitLoginForm(usr, pwd){"
    .AddNL "   var f = document.forms[0]"
    .AddNL "       f.elements['user_name'].value     = usr"
    .AddNL "       f.elements['user_password'].value = pwd"
    .AddNL "       f.submit()"
    .AddNL "}"

    wv.AddScriptToExecuteOnDocumentCreated .ToString
  End With
```

Such an "extension-function" can be added before you navigate to a certain page.

 Only the usage of this function - e.g. via:
 wv.jsRun "submitLoginForm", "myUsrID", "myPwd"

 has to be done after the document was loaded (e.g. from within Navigation_Complete or Document_Complete)

 Olaf

## #146 BooksRUs — Mar 16th, 2022, 08:51 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

The DOM is there, but you have to interact with it in js now...
 (cWebView2 is making this easy via AddScriptToExecuteOnDocumentCreated)

Code:

```
With New_c.StringBuilder

    .AddNL "function submitLoginForm(usr, pwd){"
    .AddNL "   var f = document.forms[0]"
    .AddNL "       f.elements['user_name'].value     = usr"
    .AddNL "       f.elements['user_password'].value = pwd"
    .AddNL "       f.submit()"
    .AddNL "}"

    wv.AddScriptToExecuteOnDocumentCreated .ToString
  End With
```

Such an "extension-function" can be added before you navigate to a certain page.

 Only the usage of this function - e.g. via:
 wv.jsRun "submitLoginForm", "myUsrID", "myPwd"

 has to be done after the document was loaded (e.g. from within Navigation_Complete or Document_Complete)

 Olaf

 Cool, thanks for the info. A few questions:

 1. Once I inject this function, is it "persistent" upon each subsequent Navigate event, or is it just good for *one* Navigate, then it goes away?

 2. Is there a way to "peruse" the DOM with WV? Sometimes I have to basically "dump" all of the forms and elements of a page in order to find/pick out the particular form, element or field that I need to match, fill in, check, or click.

 3. How do I "click" a button that doesn't seem to have a name?

## #147 Schmidt — Mar 17th, 2022, 06:28 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

1. Once I inject this function, is it "persistent" upon each subsequent Navigate event, or is it just good for *one* Navigate, then it goes away?

It is persistent (on that bound wv-instance) for all subsequent navigations -
 basically ensuring "your own set of global js-helper-functions"... -
 (which BTW could also be aggregated in a normal "myGlobalHelpers.js" File for better editability,
 and then added in one go like this:
 wv.AddScriptToExecuteOnDocumentCreated New_c.FSO.ReadTextContent(App.Path & "\myGlobaHelpers.js")

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

2. Is there a way to "peruse" the DOM with WV?

Sure, you can even get "visual assistance" on a loaded DOM, by opening the "Chrome-Inspector" via:
 wv.OpenDevToolsWindow

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

3. How do I "click" a button that doesn't seem to have a name?

You could learn to use "selector-string-expressions" via e.g.:
 - document.querySelector("<your selector string>") <https://developer.mozilla.org/en-US/docs/Web/API/Document/querySelector> https://developer.mozilla.org/en-US/.../querySelector
 - SomeElement.querySelector("<your selector string>") <https://developer.mozilla.org/de/docs/Web/API/Element/querySelector> https://developer.mozilla.org/de/doc.../querySelector

 Selector-Strings are extremely powerful (they allow to search for any attribute, parent/child-relation and so on...)

 At the end of the day, there's no way around doing all your "DOM-stuff" in javascript ...
 but it has to be said, that this usually requires less lines of code, compared to VB6/IE-COMbased-DOM-parsing.

 Olaf

## #148 wwolf — Mar 17th, 2022, 07:38 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

 1. Once I inject this function, is it "persistent" upon each subsequent Navigate event, or is it just good for *one* Navigate, then it goes away?

Click on different links on the Google page and watch out for the alert window:

Code:

```
Option Explicit
Private WithEvents WV As cWebView2
Private Sub Form_Activate()
    If WV Is Nothing Then
        Set WV = New_c.WebView2(Me.hWnd)
        WV.AddScriptToExecuteOnDocumentCreated "function getSiteTitle(){alert (document.title)}"
        WV.Navigate "http://www.google.com"
    End If
End Sub
Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        If Not WV Is Nothing Then WV.SyncSizeToHostWindow
    End If
End Sub
Private Sub WV_DocumentComplete()
    WV.jsRun "getSiteTitle"
End Sub
```

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

2. Is there a way to "peruse" the DOM with WV? Sometimes I have to basically "dump" all of the forms and elements of a page in order to find/pick out the particular form, element or field that I need to match, fill in, check, or click.

Code:

```
Option Explicit
Private WithEvents WV As cWebView2
Private Sub Form_Activate()
    If WV Is Nothing Then
        Set WV = New_c.WebView2(Me.hWnd)
        WV.Navigate "http://www.google.com"
    End If
End Sub
Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        If Not WV Is Nothing Then WV.SyncSizeToHostWindow
    End If
End Sub
Private Sub WV_DocumentComplete()
    Debug.Print "Source: " & WV.jsProp("Array.prototype.slice.call(document.getElementsByTagName('INPUT')).join(';')")
End Sub
```

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

3. How do I "click" a button that doesn't seem to have a name?

<https://stackoverflow.com/questions/2705583/how-to-simulate-a-click-with-javascript/2706236#2706236> https://stackoverflow.com/questions/...706236#2706236

## #149 BooksRUs — Mar 18th, 2022, 03:48 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

 The particular form I'm adding this to is the "main" form for a pretty large program that has been my project for the last 13 years. It is on the edge of having more controls than VB6 allows, so I thought perhaps having to declare WV (with events) was pushing it over the limit. I don't know if it's possible, but **it would be wonderful if we could have an ARRAY of WV (with events)**. I'm not sure if these declarations "count" towards the VB6 limits or not, but either way, this would allow an array of WV that match an array of PicBoxes that "host" the WVs that I want to use the in program (there are several). This is exactly how I'm implementing this with the Webbrowser controls that I'm trying to replace.

 I wanted to highlight the above, as I believe I have hit the VB6 maximum number of controls for a form. I was able to save the code (for a while), but then randomly, VB6 started kicking me out of the IDE without warning -- no Save, no nothing.

 I currently need 10 WebView controls to replace the webbrowser controls in the form. They are an array, so it seems to be ok with these, but I will have to see if there's anything that can be deleted from the form in order to truly replace all of them.

 I don't have a clue what would go into doing this, but I'd like to think everyone would benefit from it.

 Thanks!

## #150 jpbro — Mar 18th, 2022, 08:28 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

It would be wonderful if we could have an ARRAY of WV (with events)

The RC6.cEventCollection might be what you are looking for. I've adapted Olaf's original demo to show how you can use the cEventCollection with cWebView2 bound to a control array of PictureBoxes:

 <https://www.vbforums.com/attachment.php?s=9b84d96bf559669a1dc45b4f80d160b8&attachmentid=184364&d=1647652735> WebViewEventCollection.zip

 The code should be fairly easy to understand on it's own - it's only slightly more complicated that it would be if we the cEventCollection had an Item property, but I've made do with a separate cArrayList to hold WV references that we can access via the cEventCollection Key property). Let me know if you have any questions though.

## #151 BooksRUs — Mar 20th, 2022, 09:48 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jpbro** *[img: View Post]*

The RC6.cEventCollection might be what you are looking for. I've adapted Olaf's original demo to show how you can use the cEventCollection with cWebView2 bound to a control array of PictureBoxes:

 <https://www.vbforums.com/attachment.php?s=9b84d96bf559669a1dc45b4f80d160b8&attachmentid=184364&d=1647652735> WebViewEventCollection.zip

 The code should be fairly easy to understand on it's own - it's only slightly more complicated that it would be if we the cEventCollection had an Item property, but I've made do with a separate cArrayList to hold WV references that we can access via the cEventCollection Key property). Let me know if you have any questions though.

Cool!

 I'll check it out and let you know.

 Thanks!

## #152 Schmidt — Mar 20th, 2022, 12:09 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

I currently need 10 WebView controls to replace the webbrowser controls in the form.
 They are an array, ...

Then the easiest way to keep up with the existing "scheme" would be,
 to implement a little "Project-Private-UserControl" for your WebView2-Binding.

 E.g. something like that into a UserControl, named **ucWV**:

Code:

```
Option Explicit

Event DocumentComplete() 'redefine all the Events you need externally in this place...

Private WithEvents mWV As cWebView2, InitDone As Boolean

Public Property Get WV() As cWebView2
  EnsureBinding
  Set WV = mWV
End Property

Private Sub EnsureBinding()
  If Not Ambient.UserMode Then Exit Sub
  If mWV Is Nothing And Not InitDone And Extender.Visible Then
     InitDone = True: Set mWV = New_c.WebView2(UserControl.hWnd)
     UserControl_Resize
  End If
End Sub

Private Sub UserControl_Resize()
  If Not mWV Is Nothing Then mWV.SyncSizeToHostWindow
End Sub

'*** Event-ReDelegation of internal Events (only those which are needed externally)...
Private Sub mWV_DocumentComplete()
  RaiseEvent DocumentComplete 're-raise the internal Event to the outside
End Sub
```

After your Private UC was defined, you can now place two of them on a Form (as a Ctl-Array), then using this Code:

Code:

```
Option Explicit

Private Sub Form_Load()
  Visible = True 'always ensure Visibility of the TopLevel-Window first

  ucWV(0).WV.NavigateToString "<h1>Hello World</h1>" 'Ctl at Index 0 navigates to a String
  ucWV(1).WV.Navigate "https://google.com" 'Ctl at Index 1 navigates to Google
End Sub

Private Sub ucWV_DocumentComplete(Index As Integer) 'example for Ctl-Array-based Event-Handling
  With ucWV(Index).WV
    Debug.Print .jsProp("Array.from(document.getElementsByTagName('h1')).map(e=>e.innerText).join(',')")
  End With
End Sub
```

So it just boils down to a bit of work inside the UC (for the Event-Definition and -ReDelegation).

 Though in the end this should be quite elegant and robust.

 Olaf

## #153 BooksRUs — Mar 21st, 2022, 10:59 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Then the easiest way to keep up with the existing "scheme" would be,
 to implement a little "Project-Private-UserControl" for your WebView2-Binding.

 E.g. something like that into a UserControl, named **ucWV**:

Code:

```
Option Explicit

Event DocumentComplete() 'redefine all the Events you need externally in this place...

Private WithEvents mWV As cWebView2, InitDone As Boolean

Public Property Get WV() As cWebView2
  EnsureBinding
  Set WV = mWV
End Property

Private Sub EnsureBinding()
  If Not Ambient.UserMode Then Exit Sub
  If mWV Is Nothing And Not InitDone And Extender.Visible Then
     InitDone = True: Set mWV = New_c.WebView2(UserControl.hWnd)
     UserControl_Resize
  End If
End Sub

Private Sub UserControl_Resize()
  If Not mWV Is Nothing Then mWV.SyncSizeToHostWindow
End Sub

'*** Event-ReDelegation of internal Events (only those which are needed externally)...
Private Sub mWV_DocumentComplete()
  RaiseEvent DocumentComplete 're-raise the internal Event to the outside
End Sub
```

After your Private UC was defined, you can now place two of them on a Form (as a Ctl-Array), then using this Code:

Code:

```
Option Explicit

Private Sub Form_Load()
  Visible = True 'always ensure Visibility of the TopLevel-Window first

  ucWV(0).WV.NavigateToString "<h1>Hello World</h1>" 'Ctl at Index 0 navigates to a String
  ucWV(1).WV.Navigate "https://google.com" 'Ctl at Index 1 navigates to Google
End Sub

Private Sub ucWV_DocumentComplete(Index As Integer) 'example for Ctl-Array-based Event-Handling
  With ucWV(Index).WV
    Debug.Print .jsProp("Array.from(document.getElementsByTagName('h1')).map(e=>e.innerText).join(',')")
  End With
End Sub
```

So it just boils down to a bit of work inside the UC (for the Event-Definition and -ReDelegation).

 Though in the end this should be quite elegant and robust.

 Olaf

Ok, Newbie alert! *[img: Duck]*

 I know this will sound crazy, but I have never created my own control... but, I'm willing to give it a try.

 To start, is this an ActiveX DLL, EXE, OCX?

 For the events that have parameters, do I have to declare the same parameters in the Event declaration, then pass them "up" to the real control, or do the parameters automatically go up?

 Also, I was trying to add a "Busy" flag, as I didn't see how I can do this:
 1. Navigate somewhere
 2. wait until WV is actually there and complete
 3. continue

 So, in my main program, I basically set a global Boolean for *each* WV as WVbusy=true, just before I WV.navigate, then in the NavigationCompleted event, I reset the global WVbusy back to false.

 Since I'm "extending" the WV, I didn't know how/if I could extend it this way?

 Thanks!

## #154 saturnian — Mar 22nd, 2022, 03:53 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I programmed an ocx encapsulating the RC6's Webview2 for my own purposes. It is available** <https://www.ordoconcept.net/ordoWebview2/en/> here**, for those interested.

 It is a version which will evolve with the improvements of the RC6 and which is delivered as it is, without any guarantee. But in my software, it works perfectly. Thanks again to Olaf!

 François

## #155 avinash7 — Mar 23rd, 2022, 12:56 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

While opening the application for the 1st time after starting computer, application crashes with below details in the event viewer. From 2nd time onwards, no issues at all on same computers. This happens for both installed EXE or running project from VB6 IDE. We have kept 30 seconds timeout for BindTo method (WV.BindTo(picWV.hWnd, 30)). Timeout doesn't seem to be the issue as application crashes immediately. All computers have Evergreeen WebView2-Runtime installed which is updated automatically as well as properly registered required RC6 related DLLs in single folder. This happens on multiple computers. We are opening both EXE as well as VB6 IDE with Run as Administrator, if that matters.

 Faulting application name: VB6.EXE, version: 6.0.81.76, time stamp: 0x3592011f
 Faulting module name: ntdll.dll, version: 10.0.19041.1566, time stamp: 0xbde09443
 Exception code: 0xc0000374
 Fault offset: 0x000e6d03
 Faulting process id: 0xde4
 Faulting application start time: 0x01d83e7836746466
 Faulting application path: D:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE
 Faulting module path: C:\Windows\SYSTEM32\ntdll.dll
 Report Id: 91d860ff-e220-4c76-b631-9b18f578ba09
 Faulting package full name:
 Faulting package-relative application ID:

## #156 Schmidt — Mar 26th, 2022, 09:13 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

Ok, Newbie alert! *[img: Duck]*

 I know this will sound crazy, but I have never created my own control... but, I'm willing to give it a try.

 To start, is this an ActiveX DLL, EXE, OCX?

A "Project-Private-UserControl" is "just another module" (in your normal Std-Exe-Project)
 You add it in the same way as an additional Form (*.frm), Module (*.bas), Class (*.cls)...
 via Context-Menu in your Project-Tree as a UserControl (ending up with a *.ctl -File in your ProjectTree when saving).

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

For the events that have parameters, do I have to declare the same parameters in the Event declaration...?

Yes - and since you have to "Re-Delegate them" anyways (see the last section of my Demo),
 you will have to receive them first in the internal EventSink-Handler of the WV.

 And once you have that EventHandler selected (via DropDown),
 you can immediately copy its "Param Signature" as is - into the Event-Declaration at the Top of your little UC.

 Re-Delegating 10 Events or so, should not take more than 5 minutes of "Copy&Paste-work".

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

Also, I was trying to add a "Busy" flag, as I didn't see how I can do this:
 1. Navigate somewhere
 2. wait until WV is actually there and complete
 3. continue

Actually, you don't need to implement your own "Wait-Handling" for Navigation... since this is already built-in.
 (please look at the optional TimeOut-Param in the Navigate-Methods).

 HTH

 Olaf

## #157 Schmidt — Mar 26th, 2022, 09:23 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **avinash7** *[img: View Post]*

While opening the application for the 1st time after starting computer, application crashes

Are you sure that this crash is caused by the WebView2?

 What if you take a very simplified example (e.g. similar to the one w.wolff has posted in #148):

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
    Visible = True: Caption = "Now navigating to Google..."

    Set WV = New_c.WebView2(Me.hWnd)
        WV.Navigate "https://google.com"
End Sub

Private Sub Form_Resize()
    If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub

Private Sub WV_DocumentComplete()
    Caption = "Document complete"
End Sub
```

Put this into a virginal StdExe-Project(-Form) and compile.
 Then re-start your machine and check the little Executable, how it behaves "at first Startup, shortly after booting the machine".

 If that works, then the error is in all likelihood *not* caused by the WebView2-lib (or the wrapper).

 Olaf

## #158 BooksRUs — Mar 29th, 2022, 01:53 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

I programmed an ocx encapsulating the RC6's Webview2 for my own purposes. It is available** <https://www.ordoconcept.net/ordoWebview2/en/> here**, for those interested.

 It is a version which will evolve with the improvements of the RC6 and which is delivered as it is, without any guarantee. But in my software, it works perfectly. Thanks again to Olaf!

 François

Thanks for this!! I have downloaded and I'm trying it out. My first problem is how to set it up to call the .JSrun with Parameters? I'm converting the code from using the WebView2 directly and it's complaining:

 **Type Mismatch: array or user-defined type expected**

 I have injected a small JS with 2 parameters (user and password), so that I can auto-login to a website.

 Thanks!

## #159 BooksRUs — Mar 29th, 2022, 09:22 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

 Re-Delegating 10 Events or so, should not take more than 5 minutes of "Copy&Paste-work".

While this exercise seemed trivial enough, it appears that to provide a true "wrapper" around this with Indexes, I must re-create/declare *ALL* the other Let/Get and Methods for *each* of the properties/methods of the actual WV object.

 I did try to do this, but the .jsProp was a little confusing, while most others were straightforward. I only tried to do the ones that I'm using to change your Demo program to work with 2 of my own ucWV objects (I dynamically Load a 2nd one, just to make sure that I can dynamically Load the object).

## #160 saturnian — Apr 1st, 2022, 05:42 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

Thanks for this!! I have downloaded and I'm trying it out. My first problem is how to set it up to call the .JSrun with Parameters? I'm converting the code from using the WebView2 directly and it's complaining:

 **Type Mismatch: array or user-defined type expected**

 I have injected a small JS with 2 parameters (user and password), so that I can auto-login to a website.

Have you tried : (where FLogin is the Javascript procedure)

Code:

```
    Dim Param_Array(1) As Variant
    Param_Array(0) = "MyUserID"
    Param_Array(1) = "MyPassword"
    OrdoWebView1.jsRun "FLogin", Param_Array()
```

## #161 paliadoyo — Apr 22nd, 2022, 06:03 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I All,

 I am not able to use the RC6 WebView2 from outside VB6. In VB6 (in the IDE and compiled) everything works correctly. The same if I use the "saturnian" control (an ocx encapsulating the RC6 WebView2). The problem is when I want to use the WebView2 from, for example, PowerBuilder. I've been using RC without problems in PB for a long time, but I can't get the WebView2 to work. When I call BindTo , the PB IDE crashes (the same in a compiled version) with the same error it gives to "avinash7" (error in ntdll.dll").
 I have tested the AntView control (<https://antview.dev/downloads/> https://antview.dev/downloads/) which I believe uses the same mechanism (WebView2Loader.dll) and everything works correctly, both in VB6, PowerBuilder, even in an Excel userForm.
 If I use the saturnian control in PowerBuilder, I get the error (just instantiate the control):
 <https://www.vbforums.com/attachment.php?s=9b84d96bf559669a1dc45b4f80d160b8&attachmentid=184685> Attachment 184685
 For what it's worth, in PowerBuilder's object browser, I can see all RC6 classes except "cWebView2":
 <https://www.vbforums.com/attachment.php?s=9b84d96bf559669a1dc45b4f80d160b8&attachmentid=184686> Attachment 184686
 However, the error does not occur when creating the cWebView2 object, but when calling BindTo (or when creating it by passing the hWnd parameter).
 Any ideas?

 Thanks!

## #162 carl039 — Apr 23rd, 2022, 03:43 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

@saturnian

 Does your ActiveX control allow the setting of the userDataFolder ?
 Can't seem to find it in VB6 Object Browser.

 Thanks

## #163 saturnian — Apr 23rd, 2022, 11:39 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **carl039** *[img: View Post]*

@saturnian

 Does your ActiveX control allow the setting of the userDataFolder ?
 Can't seem to find it in VB6 Object Browser.

 Thanks

For the moment no. But if you are using a Fixed Version of WebView2, the user directory data is in the installation folder of this version (by default: C:\ProgramData\OrdoWebView2\Microsoft.WebView2.FixedVersionRuntime\Data\EBWebView)
 You can copy the contents of the C:\ProgramData\OrdoWebView2\Microsoft.WebView2.FixedVersionRuntime folder several times and assign a version of the webview2 control to each user by setting the EdgeFixedPath property.
 But this cannot be changed on the fly. You must use one OrdoWebView2 control per user, modify the EdgeFixedPath property and launch the Init method.

 Best regards

## #164 carl039 — Apr 24th, 2022, 06:01 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

For the moment no. But if you are using a Fixed Version of WebView2, the user directory data is in the installation folder of this version (by default: C:\ProgramData\OrdoWebView2\Microsoft.WebView2.FixedVersionRuntime\Data\EBWebView)
 You can copy the contents of the C:\ProgramData\OrdoWebView2\Microsoft.WebView2.FixedVersionRuntime folder several times and assign a version of the webview2 control to each user by setting the EdgeFixedPath property.
 But this cannot be changed on the fly. You must use one OrdoWebView2 control per user, modify the EdgeFixedPath property and launch the Init method.

 Best regards

 Thanks, great job on your control btw.

## #165 carl039 — Apr 24th, 2022, 06:05 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

@Olaf

 Is there any plans to include all properties as detailed <https://docs.microsoft.com/en-us/dotnet/api/microsoft.web.webview2.core.corewebview2settings.ispasswordautosaveenabled?view=webview2-dotnet-1.0.1185.39> here in particular:

 IsPasswordAutosaveEnabled
 IsGeneralAutofillEnabled

 Thanks

## #166 inspace — Jun 10th, 2022, 10:44 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

just ran across this, it's exactly what i've been looking for, so thank you.

 i have a question, hopefully someone can help.

 i've figured out mostly everything i need, but i need help with 2 things.

 1) outputting the entire html of the website into a textbox. i was able to use WV.jsProp("document.body.innerHTML") but it only captures part of the html, and i need the entire source of the page.

 2) scan entire page for elements such as ID, NAME, HREF, and display them in a textbox. basically, with the webbrowser control i would do this:

 For X = 0 To WebBrowser1.Document.All.Length
 strID = WebBrowser1.Document.All.Item(X).Id
 If Not strID = "" Then txtIDs.Text = txtIDs.Text & "(" & X & ") " & wB.Document.All.Item(X).Id & vbNewLine
 Next X

 so basically that just loops through every element on the page, checks if it has an ID, and if so it displays it in a textbox. then you can replace Id with name or href. i need to re-create that same thing here. does anyone know how to achieve this? thanks!

 ----

 edit:
 i figured out #1 i think. if anyone else is having the same issue, try using WV.jsProp("document.documentElement.outerHTML")

## #167 Schmidt — Jun 11th, 2022, 07:09 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **inspace** *[img: View Post]*

2) scan entire page for elements such as ID, NAME, HREF, and display them in a textbox.

With the WebView2 you should get more familiar with its JS-interface ...
 (doing all of the more complex stuff, by adding js-Functions beforehand - and then calling these from VB)

 For example, if you add the following js-function (via WV.AddScriptToExecuteOnDocumentCreated):

Code:

```
WV.AddScriptToExecuteOnDocumentCreated New_c.ArrayList(vbString, _
   "function getByExpression_id_tag_val_name_href(sExpr){", _
   "    var resArr = []", _
   "    for (var e of document.querySelectorAll(sExpr).values()){", _
   "        resArr.push({id:e.id, tag:e.tagName, val:e.value, name:e.name, href:e.href})", _
   "    }", _
   "    return JSON.stringify(resArr)", _
   "}").Join(vbLf)
```

You can then use this js-routine from VB (with a "Selector-Filter-Expression"-param) e.g. this way:

 Debug.Print WV.jsRun("getByExpression_id_tag_val_name_href", "[id]:not([id=''])")

 The blue marked Expression-String above, filters for:
 - all Elements which contain the ID-attribute (the leading [id] -part)
 - and of those with an id-attribute, ignore the ones which have an empty string in it (the not([id=''] -part)

 These Selector-Expressions are extremely powerful - just google for examples...

 HTH

 Olaf

## #168 NemoN — Jun 14th, 2022, 10:01 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Does anyone have an example of how to use NavigateWithWebResourceRequest with HTTP Basic Auth?

## #169 shahramdelta — Jun 22nd, 2022, 09:50 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi everyone

 I just wanted to know is there any equivalent method in webview2 for the followings:

 suppose mdocDocument is declared as an HTMLDocument

 set mhtmInput1 = mdocDocument.createElement("Input")

 mdocDocument.appendChild mhtmInput1

 set mhtmCoords = mdocDocument.getElementById("coords")

 mdocDocument.parentWindow.event

 mdocDocument_onClick

 mdocDocument.parentWindow.Blur

 Thanks in advance

## #170 wwolf — Jun 24th, 2022, 02:48 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>
 suppose mdocDocument is declared as an HTMLDocument

The HTMLDocument and the corresponding interface do not exist here. You have to program this code (a function) in JavaScript and either integrate it statically into the page or add it dynamically via VB. You can call the function from your programme at any time (also parameterised). See also: AddScriptToExecuteOnDocumentCreated, jsRun, jsRunAsync and ExecuteScript.

## #171 Meierk — Jul 25th, 2022, 07:39 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf
 The problem "Runtime Error 53: Error in FindFirstFile: ErrNum: 3, The system cannot find the path specified." persists.

 I had to clean install my windows 10, registered RC6.dll and today installed WebView2. WebView version 103.0.1264.71 installed into the folder "C:\Program Files (x86)\Microsoft\EDGE\Application", your program looks for the "EdgeWebView" folder.

 Solution? Should I rename this folder?

 Regards

## #172 Meierk — Jul 25th, 2022, 07:46 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf
 Solutions sometimes can be quick: I did what I should have done before posting.

 I copied the "Edge" folder and renamed the copy "EdgeWebView" - now my program runs smooth!

 Regards

## #173 lodep — Jul 26th, 2022, 01:03 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi !

 Thank you for this great tool !
 I am looking for a way to change the WebView2 default User-Agent.
 I know that I can send a request with NavigateWithWebResourceRequest and set the User-Agent via the headers but I need that all the resources of the page are loaded with the same User Agent. In .net the solution is to call this method :

 var settings = new Web.CoreWebView2.Settings;
 settings.UserAgent = "XXXXX";

 Is there a way to do this ?

## #174 Schmidt — Jul 26th, 2022, 03:15 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **lodep** *[img: View Post]*

I need that all the resources of the page are loaded with the same User Agent...

There's a ton of settings possible via optional commandline-string (which you can pass along in the Bind-Method) -
 "--user-agent ..." just being one of them...
 <https://peter.sh/experiments/chromium-command-line-switches/#user-agent> https://peter.sh/experiments/chromiu...es/#user-agent

 I'd try it this way first...

 HTH

 Olaf

## #175 lodep — Jul 26th, 2022, 03:46 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

There's a ton of settings possible via optional commandline-string (which you can pass along in the Bind-Method) -
 "--user-agent ..." just being one of them...
 <https://peter.sh/experiments/chromium-command-line-switches/#user-agent> https://peter.sh/experiments/chromiu...es/#user-agent

 I'd try it this way first...

 HTH

 Olaf

How can I say... JUST A HUGE **THANK YOU** !!!
 This is just amazing, it works !
 Thank you *[img: Smilie]*

## #176 lodep — Jul 27th, 2022, 03:01 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi @Olaf,

 One more question : I can't access the documenturl property sometimes (it's blank).
 I've read in previous posts that you have a workaround for this problem and that you'll need to update the DLL.
 Is it possible to access this property with a temp fix using JSPROP for example ?

 Thanks,

## #177 jenniger9 — Aug 8th, 2022, 08:30 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

can anybody share with me how to get webview2 cookies vb6

## #178 Schmidt — Aug 9th, 2022, 09:03 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jenniger9** *[img: View Post]*

can anybody share with me how to get webview2 cookies vb6

Please study (or at least search through) the postings of the discussion-thread you're currently in...
 Cookie-retrieval was already a topic here in: <https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)&p=5498591&viewfull=1#post5498591> https://www.vbforums.com/showthread....=1#post5498591

 Olaf

## #179 sergeos — Sep 5th, 2022, 12:02 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Can this tool(webview2) replace selenium to emulate user behavior on the webpage?

## #180 Schmidt — Sep 6th, 2022, 01:02 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **sergeos** *[img: View Post]*

Can this tool(webview2) replace selenium to emulate user behavior on the webpage?

In my understanding, Selenium was written as a generic Test-Tool which can "work against multiple Browser-engines".

 The WebView2 is just an encapsulation of a single Browser-engine (chromium) -
 minus the additional "generic Automation- and Test-interfaces".

 So, if you currently use Selenium for your own testing-purposes with success,
 why change it to something "less capable" and "less generic".

 Just my $0.02...

 Olaf

## #181 jangle — Sep 7th, 2022, 09:12 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello!
 Can you explain how to save downloaded web page to local disk?

## #182 sergeos — Sep 7th, 2022, 04:18 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

In my understanding, Selenium was written as a generic Test-Tool which can "work against multiple Browser-engines".

 The WebView2 is just an encapsulation of a single Browser-engine (chromium) -
 minus the additional "generic Automation- and Test-interfaces".

 So, if you currently use Selenium for your own testing-purposes with success,
 why change it to something "less capable" and "less generic".

 Just my $0.02...

 Olaf

no, I don't use selenium. but I have seen how it works. The problem is that with every update of firefox, you need to change the version of the woking executable to connect firefox to the web page.
 I also join the question above regarding use cases.

## #183 Schmidt — Sep 8th, 2022, 09:57 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jangle** *[img: View Post]*

Hello!
 Can you explain how to save downloaded web page to local disk?

Not sure, what is meant here...
 (there is an interactive right-mouseclick-context-menu in the chromium-engine,
 which allows to trigger "Save as..."-commands).

 If it's only "plain-text-" or "plain-binary"-saving of a given URL -
 then a better suited (more lightweight) helper-object is available via the winhttp 5.1 reference...

 Olaf

## #184 jangle — Sep 8th, 2022, 11:24 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Not sure, what is meant here...
 (there is an interactive right-mouseclick-context-menu in the chromium-engine,
 which allows to trigger "Save as..."-commands).

 If it's only "plain-text-" or "plain-binary"-saving of a given URL -
 then a better suited (more lightweight) helper-object is available via the winhttp 5.1 reference...
 Olaf

I have a list of about 100 web pages. I need to automatically download them and save them along with the pictures to a local disk.
 Can I use this component for this task or should I use something else?

## #185 Resurrected — Sep 19th, 2022, 10:15 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf:

 I was using RC6 version: 6.0.0.9, everything is fine.
 Then I upgraded to RC6 version: 6.0.0.10, my program broke.

 It was broken when this line of code stopped returning anything:

Code:

```
WV.jsProp("document.querySelector('input').value")
```

I have to switch back to 6.0.0.9 and my program works again.

 I notice there are some new features added to cWebView2, perhaps that's where the root of the problem is.

## #186 Schmidt — Sep 20th, 2022, 01:13 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Resurrected** *[img: View Post]*

Olaf:

 I was using RC6 version: 6.0.0.9, everything is fine.
 Then I upgraded to RC6 version: 6.0.0.10, my program broke.

 It was broken when this line of code stopped returning anything:

Code:

```
WV.jsProp("document.querySelector('input').value")
```

I have to switch back to 6.0.0.9 and my program works again.

Had to adapt the jsProp-Resolver Method internally (in 6.0.10) due to "security-reasons" -
 (the old version was using eval(...) to resolve the jsProp-expression, and that does not work in all cases).

 What the new version still supports, is "nested Prop-Expressions" - as for example:
 Debug.Print WV.jsProp("window.document.title")

 But please note, that the above "nested expression" does not contain any parentheses (which indicate function-calls).
 So, it has to be "pure properties" now (since I'm not using eval anymore for jsProp-evaluation internally).

 To solve your problem, just add a little helper-function like e.g:
 WV.AddScriptToExecuteOnDocumentCreated "function getQSValue(qsExpr){return document.querySelector(qsExpr).value}"
 ...before loading a document.

 This will allow you, to use WV.jsRun(...) instead, to get your values:
 Debug.Print WV.jsRun("getQSValue", "input")

 HTH

 Olaf

## #187 BooksRUs — Sep 21st, 2022, 01:52 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

 Re-Delegating 10 Events or so, should not take more than 5 minutes of "Copy&Paste-work".

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

While this exercise seemed trivial enough, it appears that to provide a true "wrapper" around this with Indexes, I must re-create/declare *ALL* the other Let/Get and Methods for *each* of the properties/methods of the actual WV object.

 I did try to do this, but the .jsProp was a little confusing, while most others were straightforward. I only tried to do the ones that I'm using to change your Demo program to work with 2 of my own ucWV objects (I dynamically Load a 2nd one, just to make sure that I can dynamically Load the object).

 I just wanted to BUMP this response again to see if anyone else has successfully tried/done this (making a User Control with WV so that we can have indexed UC with multiple on one form)?

 I want to replace about 17 old webbrowser controls in my VB6 program and the form has simply run out of available controls (I guess 256 is the VB6 limit). Most, if not all of them, are "hosted" on a vbalDTabControl (from vbAccelerator.com), which allows me to just set the "panel" property directly to the ucWV (user control WebView) that I created.

 The code seems to actually run fine, but only first WV created actually renders anything. The rest look like they are working correctly "under the hood" (when debugging, all events are fired), but only a blue box is rendered on the screen.

 Any help would be appreciated.

 Thanks for a great control!

## #188 Resurrected — Sep 21st, 2022, 10:46 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Had to adapt the jsProp-Resolver Method internally (in 6.0.10) due to "security-reasons" -
 (the old version was using eval(...) to resolve the jsProp-expression, and that does not work in all cases).

 What the new version still supports, is "nested Prop-Expressions" - as for example:
 Debug.Print WV.jsProp("window.document.title")

 But please note, that the above "nested expression" does not contain any parentheses (which indicate function-calls).
 So, it has to be "pure properties" now (since I'm not using eval anymore for jsProp-evaluation internally).

 To solve your problem, just add a little helper-function like e.g:
 WV.AddScriptToExecuteOnDocumentCreated "function getQSValue(qsExpr){return document.querySelector(qsExpr).value}"
 ...before loading a document.

 This will allow you, to use WV.jsRun(...) instead, to get your values:
 Debug.Print WV.jsRun("getQSValue", "input")

 HTH

 Olaf

Olaf, thanks for the reply. I will make changes to my code.

 Thanks again for your support in all those in years.

## #189 Resurrected — Sep 29th, 2022, 10:19 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf,

 Another problem related to .jsRun method.

 It seems that .jsRun returns the previous result if the current result is undefined?

 For example,

 .jsRun(js1) returns 'result of js1' if js1 works.
 .jsRun(js2) returns also 'result of js1' if js2 failed.
 .jsRun(js3) returns 'result of js3' if js3 works.

 I was assuming that .jsRun would return something like NULL when the script fails.

 Is there any possible mistake on my side? Or is there anything I don't know about .jsRun?

 Thanks

## #190 Schmidt — Sep 30th, 2022, 02:33 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Resurrected** *[img: View Post]*

It seems that .jsRun returns the previous result if the current result is undefined?

 For example,

 .jsRun(js1) returns 'result of js1' if js1 works.
 .jsRun(js2) returns also 'result of js1' if js2 failed.
 .jsRun(js3) returns 'result of js3' if js3 works.

 I was assuming that .jsRun would return something like NULL when the script fails.

I've fixed that now - returning Empty in case of a js-error (a returned js-**null**-value maps to VT_Empty in the COM-world).
 Have not compiled and re-uploaded yet, but it will come in the next release...

 In the interim you have two ways to workaround that:
 1) via your own try-catch-block within your self-defined js-Function (to return a null, or an err-string in case of an error)
 2) or relying on the existing, internal "try-catch"-mechanism of jsRun/jsProp
 .. (which in case of 2) manifests itself in: WV.jsLastError

 So, for "critical js-Calls that might fail" (and have no internal try/catch definition),
 you can always add an additional line after the WV.jsRun(...) like:
 If Len(WV.jsLastError) Then LogOrRaiseThe WV.jsLastError

 Olaf

## #191 paliadoyo — Oct 1st, 2022, 02:24 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I Olaf.
 Do you know why I can't see (and use) the cWebView2 class in Excel?
 <https://www.vbforums.com/attachment.php?s=9b84d96bf559669a1dc45b4f80d160b8&attachmentid=185944> Attachment 185944
 The same happens in other languages (powerbuilder for example), where I have had no problems using other RC6 classes

 Regards

## #192 BilalAhmed — Oct 14th, 2022, 06:17 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf,

 First of all thank you very much for a wonderful contribution!!!

 Couple of problems I am having:

 1. browser.jsProp("document.getElementById('divid').innerHTML") never returns any thing. However browser.jsProp("document.documentElement.innerHTML") returns whole document html and I then have to parse through whole document to find content which I am looking for.

 2. Below is my JS Code
 var url = 'https://abc.com/page1';
 var xhr = new XMLHttpRequest();
 xhr.open('POST', url);
 xhr.setRequestHeader('accept', 'application/json, text/plain, */*');
 xhr.setRequestHeader('Authorization', 'token');
 xhr.setRequestHeader('Content-Type', 'application/json;charset=utf-8');
 xhr.setRequestHeader('content-length', '0');
 xhr.setRequestHeader('Accept-Language', 'en-US,en;q=0.9');
 xhr.setRequestHeader('content-type', 'application/json');

 xhr.onreadystatechange = function () {
 if (xhr.readyState === 4) {
 ret = xhr.responseText; document.getElementById('dvRes').innerHtml = xhr.responseText;
 }};
 var data = '';
 xhr.send(data);

 so when I first time do browser.ExecuteScript(MyJSCode) it works well and sets output to the specified div. But, If I have to make another xhr call or even more calls, it never even attempts to make a call (I also check network calls from Dev Tools window). So in a nutshell, it makes only one xhr call and then never goes for 2nd or 3rd one.

 Any help or direct guidance will be highly appreciated!!!

 Thank you!

## #193 Schmidt — Oct 21st, 2022, 05:41 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Just a few posts above (in #186) both your questions were answered.

 Your second question "indirectly" - meaning:
 - encapsulate the remote-call in a proper js-helper-function
 ..e.g. called function RPC(URL, oData){...}
 - then add this function using: WV.AddScriptToExecuteOnDocumentCreated <your js-function-string>
 - then you can use this function via WV.jsRun("RPC", ...)

 Olaf

## #194 talatoncu — Nov 17th, 2022, 11:15 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Dear Olaf,

 I have been using Webview2-binding for a long time. Thank you very much for your valuable creation of this library.

 After I have downloaded and installed "The vbRichClient-Framework (currently at Version 6.0.10", I am in great trouble.

 Although Webview2 perfectly loads the document in the picturebox, in the NavigationCompleted event, IsSuccess is always false while WebErrorStatus=0.

 Am I doing something wrong or something has been chaned that I don't know.

 Thank you very much for your assistance.

## #195 Schmidt — Nov 19th, 2022, 09:48 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **talatoncu** *[img: View Post]*

After I have downloaded and installed "The vbRichClient-Framework (currently at Version 6.0.10", I am in great trouble.

 Although Webview2 perfectly loads the document in the picturebox, in the NavigationCompleted event, IsSuccess is always false while WebErrorStatus=0.

 Am I doing something wrong or something has been chaned that I don't know.

Have just tested this with a WV.Navigate to "https://google.com" -
 and the incoming NavigationCompleted-EventParams are: IsSuccess=True, WebErrorStatus=0

 In the MS-documentation for the IsSuccess-Param is stated:
 ...Note that WebView2 will report the navigation as 'unsuccessful' if the load for the navigation did not reach the expected completion for any reason. Such reasons include potentially catastrophic issues such network and certificate issues, but can also be the result of intended actions such as the app canceling a navigation or navigating away before the original navigation completed. Applications should not just rely on this flag, but also consider the reported WebErrorStatus to determine whether the failure is indeed catastrophic in their context.

 So, maybe the site you navigate to has an outdated certificate - or undergoes several redirects?

 Olaf

## #196 shamiur — Nov 24th, 2022, 01:48 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,

 RC6 Version 6.0.10 does not return values for html textbox on vb form. It does the same thing for WebView2Demo example as well.
 MsgBox WV.jsProp("document.getElementById('txt1').value")

 Kind Regards,

## #197 talatoncu — Nov 25th, 2022, 04:02 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Have just tested this with a WV.Navigate to "https://google.com" -
 and the incoming NavigationCompleted-EventParams are: IsSuccess=True, WebErrorStatus=0

 In the MS-documentation for the IsSuccess-Param is stated:
 ...Note that WebView2 will report the navigation as 'unsuccessful' if the load for the navigation did not reach the expected completion for any reason. Such reasons include potentially catastrophic issues such network and certificate issues, but can also be the result of intended actions such as the app canceling a navigation or navigating away before the original navigation completed. Applications should not just rely on this flag, but also consider the reported WebErrorStatus to determine whether the failure is indeed catastrophic in their context.

 So, maybe the site you navigate to has an outdated certificate - or undergoes several redirects?

 Olaf

Dear Olaf,

 Thank you very much for your valuable explanation.

 So, I think the best thing is to "...consider the reported WebErrorStatus to determine whether the failure is indeed catastrophic in their context...".

## #198 Schmidt — Nov 27th, 2022, 02:55 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **shamiur** *[img: View Post]*

RC6 Version 6.0.10 does not return values for html textbox on vb form. It does the same thing for WebView2Demo example as well.
 MsgBox WV.jsProp("document.getElementById('txt1').value")

I've explained why this happens - and also a workaround for it - in post #186 above.

 Olaf

## #199 Bob17 — Dec 1st, 2022, 10:20 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,
 is it possible that you offer some sort of error number or exception on calling BindTo method instead of just returning 0 or 1? I have a case of a shutdown of msedgewebview2.exe(s) after they initialize and after they create files/folders in the given user data folder. They just shut down very shortly after starting and the return value of BindTo is 0. I tried every suggestion from WebView2 forums without success, there is no trace of it in event logs, I have no idea how to proceed. I saw that .NET WebView2 library in such situations does return exceptions with meaningful messages, it would be great if you could offer similar information on BindTo errors, thanks.

## #200 Resurrected — Dec 9th, 2022, 09:22 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf

 I'm experiencing a problem.

 Today, out of nowhere my program suddenly ran into "couldn't initialize Webview2 binding" on startup.
 The problem remains after I tried rebooting my pc, recompliling the exe, reinstalling webview2 runtime, running in the IDE, or reresgistering RC6 binaries.

 On another PC of mine, there is no problem.

## #201 Schmidt — Dec 13th, 2022, 04:30 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Resurrected** *[img: View Post]*

Today, out of nowhere my program suddenly ran into "couldn't initialize Webview2 binding" on startup.
 The problem remains after I tried rebooting my pc, recompliling the exe,
 reinstalling webview2 runtime, running in the IDE, or reresgistering RC6 binaries.

Cannot reproduce the behaviour on my machine -
 but nevertheless have uploaded a new RC6-BaseLibs-package -
 which now contains the newest version of the MS 'WebView2Loader.dll' -
 as well as a slightly changed BindTo-Method, which deals more thouroughly with "compatible-versions"...

 Please check it out after downloading the new package (with all Dlls) -
 and re-registering RC6.dll

 @Bob17
 This might help also in your case - though not sure (it's still a "shot into the blue" -
 because - as said - cannot reproduce the behaviour).

 What helped me in this regard (in the early phases of developing it) was,
 to delete the "UserFolder" completely (in case you gave it a specific location) -
 or to change the location to a different UserFolder-Path with "full write-rights, and nothing else in it".

 HTH

 Olaf

## #202 Resurrected — Dec 14th, 2022, 07:57 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf

 I'm sorry to say that after registering the new DLLs, the problem remains the same: "couldn't initialize WebView-Binding".

 I'm running the Webview2Demo project.

 I'm starting to doubt if it's a Win7 thing. Just recently I constantly get notification from Chrome that it will not be updated in Win7 Os.

## #203 Resurrected — Dec 14th, 2022, 08:09 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

There is a fixed version Webview2 for downloading. Is it possible to test against this?

## #204 Schmidt — Dec 15th, 2022, 11:37 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Resurrected** *[img: View Post]*

There is a fixed version Webview2 for downloading. Is it possible to test against this?

Sure, the first optional Param of the BindTo-Method allows that.

 But I'd make sure, that you point it to a (lower) version of MS-Edge, which is still supported by Win7.

 Olaf

## #205 Resurrected — Dec 15th, 2022, 08:47 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf

 Good news.

 By changing the browserInstallPath of the BindTo method, I've come to the results:

 Microsoft.WebView2.FixedVersionRuntime.108.0.1462.46.x64 > couldn't initialize WebView-Binding
 Microsoft.WebView2.FixedVersionRuntime.107.0.1418.62.x64 > No errors.

 I'm running Win7 Os.

 Regards

## #206 Bob17 — Dec 20th, 2022, 06:50 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,
 Calling a vb6 COM dll that has RC6 WebView2 from a .NET application crashes the process (on calling BindTo method) with heap corruption error in ntdll.dll. This is reproducible, I have attached sample code with one c# forms project and one vb6 dll project. Could you please take a look and give some comments on it, maybe you can pinpoint the cause. Thanks!

## #207 Schmidt — Dec 20th, 2022, 08:43 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Bob17** *[img: View Post]*

Hi Olaf,
 Calling a vb6 COM dll that has RC6 WebView2 from a .NET application crashes the process...

I'm aware of that problem -
 and .NET-compiled (32Bit-x86) Host-Processes are not the only ones who cause it...
 There's also VBA (when run from within 32Bit Excel or Access.exe Processes) which has that problem.

 I have currently not investigated this deeply - but my assumption is,
 that a VB6-compiled 32Bit Host-Exe is initializing the vb6-runtime-dll differently on startup -
 (compared to an "isolatedly loaded" VB6-COM-Dll - like the RC6 on different "Host-Exes").

 In case of a VB6-compiled Host-exe the RC6-Dll "meets an already initialized vb6-runtime" (on the MainThread) -
 and in case of differently compiled Host-exes, the RC6-Dll "has to initialize the vb6-runtime itself".

 There is all sorts of other (threading-related things) to consider along with that -
 it might be that the hosting Executables where it fails, have initialized their COM(Threads) to "MTA-mode" and not "STA-Mode".

 I hesitate to invest much of my time into it - because TwinBasic offers a WebView2-Binding as well (also for VBA).
 And ".NETers" have also their own "WebView2-Binding" available in a .NET-based Class-wrapper.

 Is there a reason, why you want to use (32Bit only!)-VB6-compiled-COM-Dlls in a .NET-App?

 Olaf

## #208 Bob17 — Dec 21st, 2022, 04:49 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Well, I need to support such scenario where web browsing is in the COM dll and the caller is arbitrary application, possibly .NET-based. One interesting detail, if .NET app is based on Framework 3.5 or 4 then the app is not crashing!

## #209 Schmidt — Dec 21st, 2022, 07:04 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Bob17** *[img: View Post]*

Well, I need to support such scenario where web browsing is in the COM dll and the caller is arbitrary application, possibly .NET-based. One interesting detail, if .NET app is based on Framework 3.5 or 4 then the app is not crashing!

A VB6-Host-Exe might help (which encapsulates the whole thing - including the VB6-Form you "BindTo").

 You can even make this VB6-Executable an ActiveX-Exe-Project -
 with your current "Dll-based Public COM-Class" exposed as a Public Class of the AciveX-Exe -
 which is BTW one of the few ways, to make a VB6-Class "consumable" also by **64Bit**-Host-Processes.
 (remember, you cannot load a VB6-compiled 32Bit COM-Dll into a 64Bit-Process **directly **-
 only **Cross**-**Process**-COM-calls between 64Bit-Consumer-Process and a 32Bit-ClassHost-Process allow that feat)

 Olaf

## #210 paliadoyo — Dec 22nd, 2022, 01:52 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

I'm aware of that problem -
 and .NET-compiled (32Bit-x86) Host-Processes are not the only ones who cause it...
 There's also VBA (when run from within 32Bit Excel or Access.exe Processes) which has that problem.

 I have currently not investigated this deeply - but my assumption is,
 that a VB6-compiled 32Bit Host-Exe is initializing the vb6-runtime-dll differently on startup -
 (compared to an "isolatedly loaded" VB6-COM-Dll - like the RC6 on different "Host-Exes").

 In case of a VB6-compiled Host-exe the RC6-Dll "meets an already initialized vb6-runtime" (on the MainThread) -
 and in case of differently compiled Host-exes, the RC6-Dll "has to initialize the vb6-runtime itself".

 There is all sorts of other (threading-related things) to consider along with that -
 it might be that the hosting Executables where it fails, have initialized their COM(Threads) to "MTA-mode" and not "STA-Mode".

 I hesitate to invest much of my time into it - because TwinBasic offers a WebView2-Binding as well (also for VBA).
 And ".NETers" have also their own "WebView2-Binding" available in a .NET-based Class-wrapper.

 Is there a reason, why you want to use (32Bit only!)-VB6-compiled-COM-Dlls in a .NET-App?

 Olaf

Hello Olaf. You are right. By creating an ActiveX-Exe and exposing a cWebView2 wrapper as a public class of the ActiveX-exe I can use WebView2 in my PowerBuilder application without problems (something that directly using RC6 I have not achieved). Not ideal, but at least I have a choice.

 Thank you very much!

## #211 darjeeling — Dec 23rd, 2022, 09:01 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi,
 I have the same problem as Resurrected.
 So Resurrected seems to solve with 'Microsoft.WebView2.FixedVersionRuntime.107.0.1418.62.x64' but i have not found the installer of this version.
 Microsoft publish the latest version only.
 Anybody can help me ?

## #212 Schmidt — Dec 23rd, 2022, 03:12 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The download for fixed versions is at the bottom-right-corner of this page:
 <https://developer.microsoft.com/en-us/microsoft-edge/webview2/> https://developer.microsoft.com/en-u...edge/webview2/

 Olaf

## #213 darjeeling — Dec 24th, 2022, 02:11 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,
 Yes i see but it is a CAB.
 How to install or use this one ? I'm lost.

## #214 Schmidt — Dec 24th, 2022, 02:55 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **darjeeling** *[img: View Post]*

...but it is a CAB.

Win10 has direct support for unpacking of *.cab-Archives -
 otherwise use 7zip?

 In any case, you should end up with an unpacked Folder, named:
 ..\Microsoft.WebView2.FixedVersionRuntime.107.0.1418.62.x86
 somewhere on your local harddisk.

 Now move that whole Folder to a place you prefer in your filesystem -
 and then point the 3rd Param of the BindTo-method to that Directory-Path-(String).

 Olaf

## #215 darjeeling — Dec 24th, 2022, 03:20 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Ok Olaf,
 Thank you very much and Merry Christmas !

## #216 EasyOneX — Dec 26th, 2022, 11:07 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello.

 Is there a project / template for this so that the WebView2 can also be used as an HTML editor?

 Thank you very much and Merry Christmas!

 Andreas

## #217 Schmidt — Dec 27th, 2022, 04:47 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **EasyOneX** *[img: View Post]*

Is there a project / template for this so that the WebView2 can also be used as an HTML editor?

Sure ... did you ever try-out:
 WV.OpenDevToolsWindow
 ... after a successful WV.BindTo(...) call?

 It's (in the meantime) a full-blown IDE-environment with live-changes in the "View",
 as soon as you edit and save *.html and *.css files
 (also including support for *.js-Debugging of course).

 The DevTools-Window is highly configurable (just play around with it for a day or two) -
 and it allows write-interaction with the local FileSystem, once you specify a "trusted root-folder"
 (which is usually the root-folder of your "localhost" Webserver-instance - be that NodeJs or IIS or Nginx).

 Olaf

## #218 darjeeling — Dec 27th, 2022, 08:14 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,
 My tests:

 1 - Windows 7
 MicrosoftEdgeWebView2RuntimeInstallerX86 installed on Windows 7.

 First try:
 WV.BindTo(picWV.hWnd) => "couldn't initialize WebView-Binding"

 Second try:
 ...
 P = WV.GetMostRecentInstallPath
 (WV.GetMostRecentInstallPath seems to be the right path)
 WV.BindTo(picWV.hWnd), , P) => Crash of VB6

 and again...
 WV.BindTo(picWV.hWnd) => OK it works

 2 - Windows 10
 No problem it works at anytime !

 A solution for Windows 7 ?

## #219 Schmidt — Dec 27th, 2022, 06:42 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **darjeeling** *[img: View Post]*

P = WV.GetMostRecentInstallPath

GetMostRecentInstallPath gives the path to the "Evergreen-Version" (which the "normal" MS-installer is for).

 If you want to work against an older "Fixed-Version" (like the one in your *.cab-download),
 you have to set:

 P = "C:\some\path\to\my\unpacked\Microsoft.WebView2.FixedVersionRuntime.107.0.1418.62.x86"

 and pass that P-StringVariable into the Bind-call.

 Olaf

## #220 darjeeling — Dec 28th, 2022, 04:03 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,
 Thank you i'll try.

 Regards

## #221 Adebiyi24 — Dec 29th, 2022, 09:49 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello
 I tested the demo under w7 but I'm getting this error
 *[img: Name:  29-12-2022 15-46-44.jpg
Views: 2300
Size:  19.5 KB]*

## #222 Schmidt — Dec 29th, 2022, 10:20 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Adebiyi24** *[img: View Post]*

Hello
 I tested the demo under w7 but I'm getting this error
 *[img: Name:  29-12-2022 15-46-44.jpg
Views: 2300
Size:  19.5 KB]*

The reason for this was mentioned a few times already in this Codebank-Thread.
 (the Event-interface was changed slightly in newer RC6-versions).

 Just comment out the Event-Handler you're having problems with completely -
 and then "re-select" the EventHandler from the DropDown-ComboBox at the TopRight of the IDEs CodeEditor
 (in case you need this specific Event at all)...

 I've not yet uploaded a new Demo-Zip, because the cWebView2.cls is still getting enhancements -
 and is not yet "entirely finalized" (COM-interface-wise)...
 I'll always try my best though, to not break the stuff which is currently "in use and established"...

 Olaf

## #223 Schmidt — Jan 2nd, 2023, 05:45 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

@paliadoyo, Bob17, Resurrected
 ...or anybody else who previously had problems with loading WebView2 in "other Host-Processes" like:
 Excel(32Bit), .NET(32Bit), Powerbuilder(32Bit)

 I think I've found (and fixed) something over the Holidays regarding "more robust BindTo/Loading" ...
 Please download RC6 version 6.0.12 from the usual place and check with that (after re-registering).

 Olaf

## #224 darjeeling — Jan 2nd, 2023, 06:51 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,
 RC6 version 6.0.12:
 RC6BaseDlls.zip -> HTTP Error 404.0 - Not Found

 Regards

## #225 Schmidt — Jan 2nd, 2023, 07:55 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **darjeeling** *[img: View Post]*

Hi Olaf,
 RC6 version 6.0.12:
 RC6BaseDlls.zip -> HTTP Error 404.0 - Not Found

Oops, messed up my FTP-Upload - please try again...

 Olaf

## #226 darjeeling — Jan 2nd, 2023, 09:02 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi again,

 If WV.BindTo(picWV.hWnd) = 0 Then MsgBox "couldn't initialize WebView-Binding": Exit Sub

 Sorry but the same problem remain: "couldn't initialize WebView-Binding"

 Regards

## #227 darjeeling — Jan 2nd, 2023, 10:23 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi again,
 I tested several solutions under Windows 7 but without success.
 I confirm that I have no problem under Windows 10.
 On the other hand, the source code of Wolf (<http://www.ww-a.de/download/WebView2-Demo.zip> http://www.ww-a.de/download/WebView2-Demo.zip) works perfectly on Windows 7. However, I could not reproduce the 'BinTo' method of Wolf.
 I am disappointed.

 Regards

## #228 Schmidt — Jan 2nd, 2023, 10:59 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **darjeeling** *[img: View Post]*

...the source code of Wolf (<http://www.ww-a.de/download/WebView2-Demo.zip> http://www.ww-a.de/download/WebView2-Demo.zip) works perfectly on Windows 7.

If Wolfgangs code is working (in your VB6-**IDE **on **Win7**) -
 then you're doing something wrong in your "init-sequence"...

 Here's a minimal Demo-Code for an empty Project with a Form1 (and a reference to RC6)

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Me.Caption = New_c.Version 'just to check, which RC6-Version is in use...

  Me.Visible = True 'the FormOrPicBox.hWnd needs to be Visible, before the BindTo-call

  Set WV = New_c.WebView2
  If WV.BindTo(Me.hWnd) Then
     WV.Navigate "https://google.com"
  Else
     MsgBox "Couldn't bind the WebView to the Form.hWnd"
  End If
End Sub

Private Sub Form_Resize() 'ensure full coverage of the Containers ClientArea with the WebView
  If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub
```

HTH

 Olaf

## #229 darjeeling — Jan 3rd, 2023, 09:15 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf thank you very much for trying to help me but it still doesn't work.
 I have checked the registry keys (HKLM and HKCU) concerning the runtime and tried different versions of this runtime but nothing changes.
 I will resolve to use it under Windows 10 even if I hate Windows 10....

## #230 paliadoyo — Jan 3rd, 2023, 05:10 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

It works like a charm!
 As always, thank you very much for your interest and effort in solving our problems.

 Regards

## #231 darjeeling — Jan 6th, 2023, 09:15 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,
 Good news !
 Now it works fine on Windows 7.

Code:

```
Option Explicit
Private Cairo As cCairo
Private WithEvents WV As cWebView2

Private Sub Form_Load()
    Me.Visible = True
    Set Cairo = New_c.Cairo
    Set WV = New_c.WebView2
    Cairo.SetDPIAwareness
    If WV.BindTo(Me.hWnd) = 1 Then
        Me.Caption = "INIT OK"
        WV.Navigate "https://google.com"
    Else
        MsgBox "Couldn't bind the WebView to the Form.hWnd"
    End If
End Sub

Private Sub Form_Resize() 'ensure full coverage of the Containers ClientArea with the WebView
  If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub
```

## #232 Bob17 — Jan 9th, 2023, 06:15 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,
 I wanted to test your BindTo improvements but the vbrichclient.com is having some error 500 problems.

## #233 SearchingDataOnly — Jan 9th, 2023, 08:27 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Bob17** *[img: View Post]*

Hi Olaf,
 I wanted to test your BindTo improvements but the vbrichclient.com is having some error 500 problems.

Microsoft seems to be forcing everything related to **IE **to be removed from the browser, and I don't know if the recent frequent failures (inability to download files) of vbrichclient.com are related to Olaf's use of **IIS**.

## #234 jpbro — Jan 10th, 2023, 11:20 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

Microsoft seems to be forcing everything related to **IE **to be removed from the browser, and I don't know if the recent frequent failures (inability to download files) of vbrichclient.com are related to Olaf's use of **IIS**.

Site was just down due to mystery server reasons....it looks like it's back up now. Maybe Olaf should have used Nginx :P

## #235 SearchingDataOnly — Jan 11th, 2023, 07:16 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jpbro** *[img: View Post]*

Maybe Olaf should have used Nginx :P

Yes, Nginx + VbFcgi is much better than IIS.

## #236 saturnian — Jan 19th, 2023, 05:55 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

hello Olaf
 with currently available versions 108... and 109... of Microsoft.WebView2.FixedVersionRuntime, I have a crash that takes me out of the IDE when calling the BindTo function (WV.BindTo(picWV.hWnd, 8, EdgePath, DataEdgePath, TmpLang) ). I have the same problem when compiling the project
 I have no problem with version 103 and I have no problem with the Evergreen version either.
 Have you experienced this?
 Thank you in advance for your answer
 François

## #237 Schmidt — Jan 19th, 2023, 07:25 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

WV.BindTo(picWV.hWnd, 8, **EdgePath, DataEdgePath**, TmpLang)

On what OS do you encounter this?

 Where does EdgePath point to? (a "fixed version")?
 Where does DataEdgePath point to? (a truly writable location for this User)?

 What happens on this machine, when you call it without any "Extra-Params":
 WV.BindTo(picWV.hWnd)

 Olaf

## #238 saturnian — Jan 19th, 2023, 09:58 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>
 =Schmidt;5592174]On what OS do you encounter this?

OS is Windows 10 an Windows 11 (tested on 4 computers and on several VMs)

>

>
 Where does EdgePath point to? (a "fixed version")?

C:\ProgramData\OrdoWebView2\Microsoft.WebView2.FixedVersionRuntime

 If i use C:\Microsoft.WebView2.FixedVersionRuntime that doesn't work either

 If in these directories I put version 103.x.x.x of the Webview 2 component, it works perfectly. Newer versions 108.x.x.x and 109.x.x.x do not work

>

>
 Where does DataEdgePath point to? (a truly writable location for this User)?

C:\ProgramData\OrdoWebView2\Microsoft.WebView2.FixedVersionRuntime\Data
 or
 C:\Microsoft.WebView2.FixedVersionRuntime\Data

>

>
 What happens on this machine, when you call it without any "Extra-Params":
 WV.BindTo(picWV.hWnd)

it works perfectly because it uses the Evergreen version which works perfectly (and it's a 109.x.x.x version !!!)

 Does the latest **fixed **version 109.x.x.x available from MS work on your computers?

## #239 Schmidt — Jan 19th, 2023, 10:40 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

...the Evergreen version which works perfectly (and it's a 109.x.x.x version !!!)

At least that's then something to "automatically fall back to" (when a fixed-version-folder-param doesn't work).

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

Does the latest **fixed **version 109.x.x.x available from MS work on your computers?

Will check it out (probably this weekend) - and report back, when an easy fix was possible...

 Olaf

## #240 saturnian — Jan 19th, 2023, 12:32 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

At least that's then something to "automatically fall back to" (when a fixed-version-folder-param doesn't work).
 Olaf

the issue is that there is no trappable error.

 The BindTo function exits violently without an error message whether interpreted or compiled

 In interpreted the IDE crashes

## #241 Mojtaba — Jan 27th, 2023, 03:54 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

excellent
 I always use your tool (vbRichClient) , thank you for your hard work
 WebView2 is a very good idea
 The only problem is that I don't have permission to access and read files from the hard drive (allow file)
 For example
 In web page and javascript

HTML Code:

```
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
</head>
<body>

<script type="text/javascript">
video.load('Fire.mp4');
</script>

</body>
</html>
```

Error, this page says I receive

## #242 Schmidt — Jan 27th, 2023, 09:15 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Mojtaba** *[img: View Post]*

excellent
 ...don't have permission to access and read files from the hard drive...

The code below works for me (playing a local *.mp4 file):

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Dim VideoFile As String, sHTML As String, HtmlFile As String
      VideoFile = "c:\temp\test.mp4"
      HtmlFile = New_c.FSO.GetLocalAppDataPath & "\temp_vidplayer.html"

  sHTML = "<!DOCTYPE html><html><head><meta http-equiv='X-UA-Compatible' content='IE=edge'></head>" & _
          "<body><video controls width='100%' height='auto' src='" & VideoFile & "'/></body></html>"
  New_c.FSO.WriteTextContent HtmlFile, sHTML, True

  Me.Visible = True 'ensure Visibility before the Bind-call
  Set WV = New_c.WebView2(Me.hWnd)
      WV.Navigate HtmlFile
End Sub

Private Sub Form_Resize()
  If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub
```

Olaf

## #243 saturnian — Jan 28th, 2023, 03:33 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

hello Olaf
 have you had time to investigate the issue with the latest fixed versions of webview2?
 Thanks in advance
 Best regards
 François

## #244 Schmidt — Jan 28th, 2023, 11:30 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

...have you had time to investigate the issue with the latest fixed versions of webview2?

Yes, have done that just now ...
 and I'm not able to reproduce the problems (here on a Win10 with all Updates).

 Used the following relevant code-lines.

Code:

```
  P = "c:\temp\Microsoft.WebView2.FixedVersionRuntime.109.0.1518.70.x86"

  If WV.BindTo(picWV.hWnd, 8, P, "c:\temp\WV2UserData") = 0 Then
     MsgBox "couldn't initialize WebView-Binding": Exit Sub
  End If
```

Like you, I've used an explicitly given "fixed-runtime-path" as well as a "local UserData-Path"...

 What I took care of was the following:
 - I've downloaded the **x86**-flagged, fixed runtime from the MS-site (it comes in a *.cab-file)
 - FWIW, I've unpacked this *.cab with 7zip (and not the built-in Win10-CabViewer)
 - placed the unpacked Folder in a Parent-Folder I have **full write-rights** on ("c:\temp\...")
 - created a new, empty UserDirectory in "c:\temp\WV2UserData" (before trying the Bind-call)

 Things to note (things of importance, regarding the error on your side):
 - already mentioned, but make sure to use the 32bit-(x86)version of the fixed-runtime
 - my "User-Directory" is not a sub-directory of the "fixed-runtime-folder"
 - neither Directory is causing "potential File-Virtualization-issues" (due to placement in the C:\Programs\-Folder)
 - neither Directory is provocing "treatment as an x64-Program" (due to placement in the C:\Programs\-Folder)
 - Me.Caption = New_c.Version ...shows "6.0.12" (just make sure to report this "live", to rule out regfree-loading of older versions)

 HTH

 Olaf

## #245 saturnian — Jan 29th, 2023, 09:52 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

 What I took care of was the following:
 - I've downloaded the **x86**-flagged, fixed runtime from the MS-site (it comes in a *.cab-file)
 - FWIW, I've unpacked this *.cab with 7zip (and not the built-in Win10-CabViewer)
 - placed the unpacked Folder in a Parent-Folder I have **full write-rights** on ("c:\temp\...")
 - created a new, empty UserDirectory in "c:\temp\WV2UserData" (before trying the Bind-call)

 Things to note (things of importance, regarding the error on your side):
 - already mentioned, but make sure to use the 32bit-(x86)version of the fixed-runtime
 - my "User-Directory" is not a sub-directory of the "fixed-runtime-folder"
 - neither Directory is causing "potential File-Virtualization-issues" (due to placement in the C:\Programs\-Folder)
 - neither Directory is provocing "treatment as an x64-Program" (due to placement in the C:\Programs\-Folder)
 - Me.Caption = New_c.Version ...shows "6.0.12" (just make sure to report this "live", to rule out regfree-loading of older versions)

 HTH

 Olaf

I checked all the previous items except:
 Me.Caption = New_c.Version ...shows... ... ... "6.0.10" !

 Because, on the site <http://vbrichclient.com/> http://vbrichclient.com/ it is displayed:

>

>

 Downloads
 The Base-Dlls of the current toolset (version 6) are contained in this minimum-package:
 <http://vbrichclient.com/downloads/RC6BaseDlls.zip> RC6BaseDlls.zip (~ 3.5MB, **current version: 6.0.10**, **last updated: 2022-07-31** ... including SQLite-version: 3.39.2)

so I didn't think to download RC6BaseDlls.zip again ! I remained convinced that the current version was 6.0.10 !
 I won't be fooled now ! *[img: Wink]*

 Version 6.0.12 fixes all issues ! Thanks

 François

## #246 astronald — Feb 24th, 2023, 10:53 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi have nice day thanks for this complemente,
 i have a simple question
 i am try navigate to <http://jsfiddle.net/qqzxtk67/> http://jsfiddle.net/qqzxtk67/
 this fire new window
 How i can manipulate this thanks?

## #247 Schmidt — Feb 25th, 2023, 06:26 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **astronald** *[img: View Post]*

i am try navigate to <http://jsfiddle.net/qqzxtk67/> http://jsfiddle.net/qqzxtk67/
 this fire new window
 How i can manipulate this thanks?

Not sure, what you mean exactly (with "manipulate this")...

 What do you want to achieve in the end?

 Olaf

## #248 tmighty2 — Feb 25th, 2023, 02:48 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Can your control make use of OnFrameLoad event?
 I need to inject JS at this point of time.

## #249 Schmidt — Feb 26th, 2023, 12:14 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

Can your control make use of OnFrameLoad event?
 I need to inject JS at this point of time.

Have not played around with iFrames so far (in the context of WV2 and the stuff i use it for currently) -
 but the documentation says (about the AddScriptToExecuteOnDocumentCreated-method):

 The injected script will apply to all future top level document **and **child frame navigations...

 With that, you should be able to place dedicated js-functions and -variables (for use from within an iFrame),
 without the need of "waiting for a FrameLoad-event".

 If you're already using AddScriptToExecuteOnDocumentCreated -
 can you post a short example, what you want to achieve - and where the problem comes in?

 Olaf

## #250 tmighty2 — Feb 26th, 2023, 05:07 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I need to inject scripts to change the pages colors for handicapped users, and this is the only event where I can do it.

## #251 Schmidt — Feb 26th, 2023, 08:00 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

I need to inject scripts to change the pages colors for handicapped users, and this is the only event where I can do it.

I'm quite sure, that AddScriptToExecuteOnDocumentCreated can accomplish what you want
 (without using that event).

 Please post some code-example.

 Olaf

## #252 softv — Feb 27th, 2023, 09:06 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Dear Olaf,

 First of all, as written by me quite a few times in our forums in the past, I remain in awe of both your fabulous work and your fabulous heart. I am happy that I finally got time to dig into your vbRichClient-Framework, learning many aspects of it, slowly and steadily, since the past 2-3 weeks.

 Thanks a TON to you, I learnt the "fundamentals" of 'Open Source' licenses for the first time, from your '_Library-Licenses.txt' file. **My Hearty Thanks in TONs**, time and again, to you and all the developers of all the "Open Source" libraries/tools.

 Just thought of letting you know that some of the links (to the 'Open Source' libraries) in the aforesaid text file are not working now. For e.g., for LZ4, "http://cyan4973.github.io/lz4/" does not work now. As far as I have explored, "https://github.com/lz4/lz4" looks to be the correct page now. Similarly, "https://zlib.net/zlib_license.html" works for the latest license page of 'zlib' (and not <http://www.gzip.org/zlib/zlib_license.html> http://www.gzip.org/zlib/zlib_license.html). You may anyway double-check on this from your side too because I am quite new to using "Open Source" libraries. I just felt that I should let you know my findings so that you can replace the non-working links with the right ones in the '_Library-Licenses.txt' file in the next release of vbRichClient-Framework, if at all my findings are correct. If my findings are not correct, kindly bear with me.

 Well, the **WebView2Loader.dll** found in your latest "RC6BaseDlls.zip" package is 112KB (115, 128 bytes) with version number of **1.0.1462.37**.

 Specifically, with respect to the deployment of one's compiled application in a different system
 --
 Should I always use this '1.0.1462.37' version numbered WebView2Loader.dll only? (Presuming your cWebView2 makes calls specific to this DLL only)
 (OR)
 **Can I** use the latest WebView2Loader.dll also? The version number of the latest WebView2Loader.dll is "1.0.1587.40" as of today, as per <https://www.nuget.org/packages/Microsoft.Web.WebView2> https://www.nuget.org/packages/Microsoft.Web.WebView2, if I am right.
 (OR)
 **Should I always use the latest WebView2Loader.dll only? **
 --

 Thanks in TONs once again. I sometimes go silent thinking of the immense magnitude/depth of your work for this world society. **God Bless you, olaf! God Bless All!**

 Kind Regards.

## #253 softv — Feb 28th, 2023, 10:55 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Dear Olaf,

 I noticed in <https://www.cairographics.org/news/> https://www.cairographics.org/news/ that a new release of Cairo library has happened after quite a long gap of time. This news got me really excited thinking that this would mean that "perhaps" you have now started working on releasing the next RC6 build covering all the new features of Cairo library. If so, then my wishes in advance to you for the new/updated release of RC6.

 Kind Regards.

## #254 softv — Feb 28th, 2023, 11:10 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

And, while I await with patience your answer to my query regarding 'WebView2Loader.dll' in <https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)&p=5596777&viewfull=1#post5596777> post no. 252, dear Olaf,

 Coming to 'Cairo', I have a small query. How to **include** **blur** **also to a** **text with shadow**? For instance, in CSS, the 'text-shadow' property has 'blur radius' too (ref: <https://developer.mozilla.org/en-US/docs/Web/CSS/text-shadow> https://developer.mozilla.org/en-US/...SS/text-shadow)

 I understood by reading some materials in the net that Cairo does not provide the 'blur' feature on text shadows straightaway but it can be achieved by writing a suitable routine in c++. I am not comfortable in c++ coding though.

 I do remember reading (a week back or so) about '**Gaussian blur**' in your 'Cairo' tutorials AND/OR in some of your forum messages but I saw 'Gaussian blur' applied to images only, if my understanding was right.

 Honestly, I have not found time yet to go through all your **excellent** 'Cairo' tutorials. But, I do have "sincerely" gone through some of them and learnt a good lot about text rendering on paths. If indeed 'blur on text shadow' is already covered by you in any of your tutorials OR in any of your forum messages, then kindly intimate me the same and I shall go through them. Or, if such tutorials/messages do not exist yet, then, if at all and when your time permits, **kindly provide me a code snippet** which would "exactly" mimic what the following CSS would achieve when applied for any text.
 --
 "text-shadow: h-offset v-offset blur-radius color;"
 --

 I take this opportunity to once again thank you for both your work and your heart. God Bless you! God Bless ALL!

 Kind Regards.

## #255 astronald — Feb 28th, 2023, 02:30 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Not sure, what you mean exactly (with "manipulate this")...

 What do you want to achieve in the end?

 Olaf

Thanks i want Execute javascript on new windows and read elements

## #256 Schmidt — Mar 1st, 2023, 06:41 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **softv** *[img: View Post]*

**Can I** use the latest WebView2Loader.dll also?

You can, but you don't have to...

 WebView2Loader.dll is just a small "communication-vehicle" between your App -
 and the "rolling, evergreen" Edge/Chromium-releases MS performs under the covers of your OS automatically.

 If there's a new feature I expose in new RC6.cWebView2-releases
 (using potentially "new, extended communication-interfaces" of the WebView2Loader),
 then I'll deliver the RC6-bundle with that new(er) WebView2Loader.dll as well, of course.

 Thanks for the FOSS-Link-updates - I've included them now in _LibraryLicenses.txt.

 As for "plain cairo-questions" ... it's better to ask them in the normal Forum in the future -
 (to keep this thread here related to the RC6.cWebView2-class and its usage).

 But here's a little helper-function for "shadowed text-output":

Code:

```
Option Explicit

Private Sub Form_Load()
  Dim CC As cCairoContext
  Set CC = Cairo.CreateSurface(800, 600).CreateContext 'create the "main"-Surface-context

  CC.Paint 1, Cairo.CreateSolidPatternLng(&H333333) 'ensures a dark-grey background

  CC.SelectFont "Times New Roman", 22, vbGreen, False, True
  DrawTextWithShadowOn CC, 10, 10, 300, 50, "Some shadowed Text", 4, vbYellow, 1, 1

  Set Picture = CC.Surface.Picture
End Sub

Sub DrawTextWithShadowOn(CC As cCairoContext, x, y, dx, dy, Text As String, _
    Optional ByVal R& = 3, Optional ByVal ShadowColor&, Optional ByVal xOffs#, Optional ByVal yOffs#)
    CC.Save
      CC.TranslateDrawings x, y

      CC.PushGroup CAIRO_CONTENT_COLOR_ALPHA, 0, 0, dx, dy
        CC.DrawText 0, 0, dx, dy, Text, True 'render the text within a Push/Pop-sequence
      Dim Pat As cCairoPattern: Set Pat = CC.PopGroup(True) 'retrieve the output via Pop (in a Pattern-Obj)

      With Cairo.CreateSurface(dx, dy).CreateContext 'create a temp-surface in the dimensions of the Text-Rect
        .Paint 1, Pat '"play back" the pattern from the above "Push/Pop-Operation" on this temp-Surface

        'and what follows now, is simply two calls (1. rendering the blurred-temp-surface, and 2. rendering the unblurred srf on top)
         CC.RenderSurfaceContent .Surface.GaussianBlur(R / 2, R, True, ShadowColor), xOffs - R, yOffs - R
         CC.RenderSurfaceContent .Surface, 0, 0
      End With
    CC.Restore
End Sub
```

As for upgrading to new(er) Cairo-Releases... the current Cairo.Version 1.17.02
 (which includes some "cherry-picks" from later releases),
 it seems to work nicely and stable (without mem-leaks) - so unless there's some new features which are a "must have",
 I'm not planning on updating to later versions anytime soon.

 Olaf

## #257 softv — Mar 1st, 2023, 08:34 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

You can, but you don't have to...
 ... .. .

 But here's a little helper-function for "shadowed text-output":
 .. .. .
 Olaf

Thanks a LOT, Olaf. As ever, you are so helpful with all the detailed information regarding WebView2Loader.dll, etc.

 The helper function worked like a charm to create "text shadows with blur"! The best way to see its beauty in action was to change the various parameters (color, offsets, etc.) and see the corresponding effects. It was wonderful. Really. I will now try to take time to understand more of Gaussian Blur.

 And, I am **really sorry** about posting my cairo-related query in this "RC6 cWebView2"-related thread. Only 5 minutes after posting it, I suddenly realised my mistake. I **will be careful in the future**. Sorry again.

 Kindest Regards.

## #258 tmighty2 — Mar 4th, 2023, 03:48 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I have uploaded the script here:
 <https://drive.google.com/file/d/1ovskmmC8uNatfeLLRNDZOw8riURWlSxb/view?usp=sharing> https://drive.google.com/file/d/1ovs...ew?usp=sharing

 When you inject the entire text and then navigate to web.whatsapp.com, the selected chat will be blue.
 That is how you can test if it worked.

 Thank you!

## #259 tmighty2 — Mar 4th, 2023, 06:36 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I have tried it like this:

Code:

```
Private Sub Form_Load()
  Visible = True '<- it's important, that the hosting TopLevel-Form is visible...
                 '...(and thus the Child-PicBox indirectly as well) - before we Bind the PicBox to the WebView

  Set WV = New_c.WebView2 'create the instance
  If WV.BindTo(picWV.hWnd) = 0 Then MsgBox "couldn't initialize WebView-Binding": Exit Sub

'  Set WV = New_c.WebView2(picWV.hWnd) 'create the instance
'  If WV Is Nothing Then MsgBox "couldn't initialize WebView-Binding": Exit Sub
  LocalWebViewInit 'initialize the WebView for local usage here in our Form

  Dim sWA$
  sWA = GetText("c:\users\myuser\desktop\whatsapp.txt")

  WV.AddScriptToExecuteOnDocumentCreated sWA

End Sub

Private Sub cmdNavigate_Click()

   WV.Navigate "https://web.whatsapp.com"

End Sub

Private Function GetText(ByVal u As String) As String

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim strm As Scripting.TextStream
    With fso
        Set strm = .OpenTextFile(u, ForReading, False, TristateTrue)
        Do Until strm.AtEndOfStream
            GetText = strm.ReadAll
        Loop
        strm.Close
    End With

    strm.Close

End Function
```

But that doesn't work. The chat should highlighted with blue color, but it's not.

 Also, calling the JS functions that I added does not work.

## #260 Schmidt — Mar 4th, 2023, 08:18 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

...calling the JS functions that I added does not work.

There is an error in your larger js-Code-File:

Code:

```
const setupWhatsapp = () => {
   alert('before highlighter');
   highlighter.initHighlighting();
   alert('after highlighter);  // <- missing string-quote-char
  overrideSelectedItemBackgroundColor(); // <- therefore, this line will not execute
};
```

BTW, do you get any of the "other alerts" you've placed there?

 You could also place this little EventHandler-snippet via AddScriptOnDocumentCreated (in case you know the iFrame-ID):

Code:

```
document.getElementById('my_iframe_id').onload = function() {
    // load your extra-modules (including your inits and color-replacements) here
}
```

Cannot further play around with this myself, because I have no "whatsapp"-account.

 Olaf

## #261 Schmidt — Mar 4th, 2023, 09:41 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **astronald** *[img: View Post]*

...i want Execute javascript on new windows and read elements

Not sure, whether the following example does what you had in mind -
 but the demo below will suppress the "WV2-built-in default-BrowserWindow-Popups",
 and will simply create additional VB-Form-instances instead (each using their own internal WV-Instance of course).

Code:

```
Option Explicit

Public WithEvents WV As cWebView2, WithEvents tmrNewWindow As cTimer

Private Sub Form_Load()
  Me.Visible = True 'ensure Visibility before the Bind-call

  Set WV = New_c.WebView2(Me.hWnd) 'does an implicit bind-call (within this constructor-method)
      WV.AddScriptToExecuteOnDocumentCreated "function openPopup(uri){window.open(uri,'','width=800,height=600')}"

  If Forms.Count = 1 Then 'it's the first instance of this Form-Type (so we show an entry-demo-page, based on our own html-text)
     Const btnDef$ = "<input value='show new window' type='button' on" & "click='openPopup(""https://google.com"")'>"
     WV.NavigateToString btnDef
  End If
  WV.SyncSizeToHostWindow
End Sub

Private Sub Form_Resize()
  If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub

Private  Sub WV_NewWindowRequested(ByVal IsUserInitiated As Boolean, IsHandled  As Boolean, ByVal URI As String, NewWindowFeatures As RC6.cCollection)
  IsHandled = True '<- suppress the Browser-Controls default-handler
  NewWindowFeatures.Prop("URI") = URI 'let's enhance the features-(JSON)collection about the URI-Property
   Set tmrNewWindow = New_c.Timer(10, True, NewWindowFeatures) 'and now  buffer the features-(JSON)collection in the timer-tag, to be able to  leave here as soon as possible for the sake of "non-blocking"
End Sub

Private Sub tmrNewWindow_Timer() 'this decoupling oneshot-timer only fires in case of a new Window-Request
  ShowNewWindow tmrNewWindow.Tag 'pass the JSON-collection along (from the timers Variant-Typed Tag-Property)
End Sub

Public Sub ShowNewWindow(NewWindowFeatures As cCollection)
  Set tmrNewWindow = Nothing 'we can safely destroy the "one-shot"-timerInstance now (since it delivered its Tag-Data)

  Dim F As New Form1
      F.Show , Me  'let's show the new form (after setting its feature-Collection above)

  With NewWindowFeatures
     If .Prop("HasPosition") Then F.Move .Prop("Left") * Screen.TwipsPerPixelX, _
                                         .Prop("Top") * Screen.TwipsPerPixelY
     If .Prop("HasSize") Then F.Move F.Left, F.Top, .Prop("Width") * Screen.TwipsPerPixelX, _
                                                    .Prop("Height") * Screen.TwipsPerPixelY
     F.Caption = .Prop("URI")
     F.WV.Navigate .Prop("URI")

     Debug.Print .SerializeToJSONString '(just to show, what else is contained in the features-collection)
  End With
End Sub
```

HTH

 Olaf

## #262 tmighty2 — Mar 4th, 2023, 01:39 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thank you, Olaf.

## #263 tmighty2 — Mar 4th, 2023, 05:28 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

How could I call this js function using WV2?

 Dim c As cStringBuilder
 Set c = New cStringBuilder

 c.Append "function GetScrollPosX()"
 c.Append "{"
 c.Append " return (Math.max(document.body.scrollLeft,document.documentElement.scrollLeft));"
 c.Append "}"

## #264 tmighty2 — Mar 4th, 2023, 05:30 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

And I have to say Congralatutions and thank you. I use your dhRichClient in all my apps. I could not live without it. And I hate .NET. The IDE is so slow that I just can't use it. Really, I need 5x as much time just because it responds so slowly and has such bad Edit and Continue capabilities. I am so much quicker with VB6. I use .NET for simpler projects only or for Unity which itself is slow to work with. But for quick typing, nothing is better than VB6.

## #265 Schmidt — Mar 4th, 2023, 08:22 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

How could I call this js function using WV2?

 Dim c As cStringBuilder
 Set c = New cStringBuilder

 c.Append "function GetScrollPosX()"
 c.Append "{"
 c.Append " return (Math.max(document.body.scrollLeft,document.documentElement.scrollLeft));"
 c.Append "}"

I'd use the StringBuilder via New_c constructor (for better "regfree-mode-compatibility"):

Code:

```
  With New_c.StringBuilder
      .AddNL "function GetScrollPosX(){"
      .AddNL "  return Math.max(document.body.scrollLeft,document.documentElement.scrollLeft)"
      .AddNL "}"

      WV.AddScriptToExecuteOnDocumentCreated .ToString
  End With
```

So, the adding of this function-code-snippet happens in the line before the "End With"-destruction of the StringBuilder.

 The functioin-call itself can then be done from the VB6-side by specifying the funcName (without parens) in WV.jsRun:
 Debug.Print WV.jsRun("GetScrollPosX") 'optional Parameter-Passing via ParamArray is possible

 And JFTR, also "pure-property-access" is possible, without any "extra-function-definition" -
 (even for "nested-props", as long as there's "only dots and no parentheses" in the prop-hierarchy-string):
 Debug.Print WV.jsProp("document.body.scrollLeft")

 Olaf

## #266 tmighty2 — Mar 4th, 2023, 09:10 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

How do I get the currently shown URL, please?

## #267 tmighty2 — Mar 4th, 2023, 09:47 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Is there perhaps a way to include an ad blocker?

## #268 tmighty2 — Mar 4th, 2023, 10:07 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I am trying to run this JS code to scroll the window down:

 WV.jsRunAsync "window.scrollBy({ top: 300, behavior: "smooth" });"

 It does not work. What am I doing wrong?

## #269 Schmidt — Mar 5th, 2023, 01:02 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

WV.jsRunAsync "window.scrollBy({ top: 300, behavior: "smooth" });"

 It does not work. What am I doing wrong?

What you're doing wrong is, to not read the Parameter-Names in the Intellisense-Popups... *[img: Wink]*

 Both, WV.jsRun and WV.jsRunAsync expect a "function-name" in the first argument.

 So, the easiest way is, to "wrap the little code-block" in your own user-function - as e.g.:

 WV.AddScriptToExecuteOnDocumentCreated "function scrollSmoothlyBy(x,y){ window.scrollBy({left:x, top:y, behavior:'smooth'}) }"

 You can then later call it e.g. this way (passing the x,y arguments after the function-name):
 WV.jsRunAsync "scrollSmoothlyBy", 5, 15

 As for extensions like AdBlock... those were not supported the last time I've looked...
 Here's a posting on stack-overflow: <https://stackoverflow.com/questions/61262830/how-to-enable-extension-on-webview2> https://stackoverflow.com/questions/...on-on-webview2

 Olaf

## #270 tmighty2 — Mar 5th, 2023, 04:42 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Can I get the current URL somehow?

## #271 tmighty2 — Mar 5th, 2023, 04:54 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Would you recommend to add all the functions that I need for example at program start up, correct?

## #272 tmighty2 — Mar 5th, 2023, 05:14 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I am adding this code via "wv.AddSCriptToExecuteOnDocumentCreated" when initiating WV:

Code:

```
   function pSetFocus()
   {
        alert('focus');
      var element = document.querySelector("*[autofocus]");
      if (element)
      {

      }
      else
      {
         element = document.querySelector("form input:not([type=hidden])");
      }

      if (element)
      {

      }
      else
      {
         element = document.querySelector("input#search");
      }
      if (element)
      {

      }
      else
      {
         element = document.querySelector("input#search.ytd-searchbox");
      }
      if (element)
      {

      }
      else
      {
      }

     if (element)
     {
         var sType=element.type;
         if (sType != "email")
        {
             element.selectionStart = element.selectionEnd = element.value.length;
        }
         element.focus();
         var y = element.getBoundingClientRect().top + window.scrollY;
         window.scroll(
         {
             top: y,
             behavior: "smooth"
         });
     }
```

Then, when I want to set the focus to the first input element, I call:

Code:

```
    WV.SetFocus FocusReason_PROGRAMMATIC
    WV.jsRunAsync "pSetFocus();"
```

It doesn't have any effect.
 What am I doing wrong?

## #273 Niya — Mar 5th, 2023, 08:27 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

Then, when I want to set the focus to the first input element, I call:

Code:

```
    WV.SetFocus FocusReason_PROGRAMMATIC
    WV.jsRunAsync "pSetFocus();"
```

It doesn't have any effect.
 What am I doing wrong?

Olaf just said that it expects a function name so I'd expect that your code should be this instead:-

Code:

```
WV.jsRunAsync "pSetFocus"
```

## #274 tmighty2 — Mar 5th, 2023, 08:46 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I have tried that as well, it doesn't work.

## #275 Niya — Mar 5th, 2023, 09:08 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Ah well that's all I got. I have no experience using WebView. Perhaps someone else can try their hand at the problem.

## #276 tmighty2 — Mar 5th, 2023, 09:10 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I was stupid. My JS code was faulty. I've been using a different browser before, and it was more foregiving.
 Once I fixed it, it went fine!

## #277 Schmidt — Mar 6th, 2023, 12:48 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

Would you recommend to add all the functions that I need for example at program start up, correct?

Of course... and you don't have to call: WV.AddScriptToExecuteOnDocumentCreated
 for every little function-snippet separately (after the Bind-call was successful).

 Easier would be, when you place all your helper-functions in a single "MyHelpers.js"-file...
 and then use an UTF8-capable TextFile-ReadHelper to load them in a single line of code:

 WV.AddSCriptToExecuteOnDocumentCreated New_c.FSO.ReadTextContent("MyHelpers.js", False, CP_UTF8)

 And when you use a decent TextEditor like VSCode to edit this js-File, then you'd also produce less "basic typos with js-syntax",
 since this Editor will give you nice "visual hints" about "missing parentheses" or "missing quotechars around strings".

 As for "how to get the current Document-Title or Document-URL" -
 there's already two Properties in place on the cWebView2-class:
 - WV.DocumentURL
 - WV.DocumentTitle

 Please do some "basic investigations" on the library (RC6) or current Class (RC.cWebView2) first,
 before you start working on something concrete... (e.g.via ObjectExplorer, reachable via <F2>-Key).

 There's only about 50 Methods, Properties and Events behind the WV2-class -
 (and half of them are "IsThatEnabled" or "AreTheseAllowed" basic boolean-Props).

 Olaf

## #278 xiaoyao — Mar 6th, 2023, 04:59 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

GetEdgeWebViewInfo：
 ---------------------------
 Microsoft Edge WebView2 Runtime，Version:110.0.1587.63
 ---------------------------

Code:

```
Function GetEdgeWebViewInfo(Optional CheckX64 As Boolean) As String
   On Error Resume Next
   Dim WshShell As Object
   Dim WebView2_Version As String, WebView2_Name As String
   Dim RegPath As String
   Set WshShell = CreateObject("Wscript.Shell")

    RegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(CheckX64, "WOW6432Node\", "") & "Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}\"
    'InputBox "", "", RegPath
    WebView2_Name = WshShell.RegRead(RegPath & "name")
    WebView2_Version = WshShell.RegRead(RegPath & "PV")
    GetEdgeWebViewInfo = WebView2_Name & IIf(WebView2_Version <> "", "，Version:" & WebView2_Version, "")
End Function
```

## #279 xiaoyao — Mar 6th, 2023, 05:50 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Not sure, whether the following example does what you had in mind -
 but the demo below will suppress the "WV2-built-in default-BrowserWindow-Popups",
 and will simply create additional VB-Form-instances instead (each using their own internal WV-Instance of course).

Code:

```
Option Explicit

Public WithEvents WV As cWebView2, WithEvents tmrNewWindow As cTimer

Public Sub ShowNewWindow(NewWindowFeatures As cCollection)
  Set tmrNewWindow = Nothing 'we can safely destroy the "one-shot"-timerInstance now (since it delivered its Tag-Data)
  *
     F.Caption = .Prop("URI")
     F.WV.Navigate .Prop("URI")

     Debug.Print .SerializeToJSONString '(just to show, what else is contained in the features-collection)
  End With
End Sub
```

HTH

 Olaf

if new webview by post, how to do?not only by url
 F.WV.Navigate .Prop("URI")
 maybe need like webbrowser
 postdata="user=abcd&pass=aaabbb"
 httpheadStr="ref=www.baidu.com" & vbcrlf & "header2key=333"

 webbrowser1.Navigate url,httpheadStr,postdata
 in java,new webview only need input webview ID,IT WILL AUTO DO POSTdada or http header info.

Code:

```
Function WebPost(url As String, PostInfo As String, Web1 As WebBrowser, Optional ref As String, Optional cookie As String)
'[mycode_id:1274],edittime:2011-5-10 上午 12:26:45
On Error Resume Next
ReDim AByte(0) As Byte
PackBytes AByte(), PostInfo

      Dim vPost As Variant

      vPost = AByte ' Assign the byte array to a VARIANT

      Dim vHeaders As Variant

      vHeaders = "Content-Type: application/x-www-form-urlencoded" + Chr(10) + Chr(13)
      If ref <> "" Then vHeaders = vHeaders & "Referer:" + ref & vbCrLf
      If cookie <> "" Then vHeaders = vHeaders & "Cookie:" + cookie & vbCrLf
      Web1.Navigate url, , , vPost, vHeaders
End Function
```

## #280 xiaoyao — Mar 6th, 2023, 06:19 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Win10 has direct support for unpacking of *.cab-Archives -
 otherwise use 7zip?

 In any case, you should end up with an unpacked Folder, named:
 ..\Microsoft.WebView2.FixedVersionRuntime.107.0.1418.62.x86
 somewhere on your local harddisk.

 Now move that whole Folder to a place you prefer in your filesystem -
 and then point the 3rd Param of the BindTo-method to that Directory-Path-(String).

 Olaf

Code:

```

OK = Web1.BindTo(Me.hWnd, , App.Path & "\WebviewSdk\Microsoft.WebView2.FixedVersionRuntime.109.0.1518.78.x86") <> 0
'OK = Web1.BindTo(Me.hWnd, , App.Path & "\WebviewSdk\Microsoft.WebView2.FixedVersionRuntime.110.0.1587.63.x86") <> 0

msgbox "web1.UserAgent=" & WV.UserAgent
msgbox "navigator.appVersion=" & Web1.jsProp("navigator.appVersion")
```

 view webview version online:

 <https://ie.icoa.cn/> https://ie.icoa.cn/

 ====================
 Microsoft.WebView2.FixedVersionRuntime.109.0.1518.78.x86.cab

 内核版本 (Version) WebKit 537.36 Chrome 109.0.0.0
 完整代码 (UA)
 Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36 Edg/109.0.1518.78
 ====================
 Microsoft.WebView2.FixedVersionRuntime.110.0.1587.63.x86.cab:
 内核版本 (Version) WebKit 537.36 Chrome 110.0.0.0
 完整代码 (UA)
 Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36 Edg/110.0.1587.63

## #281 xiaoyao — Mar 6th, 2023, 06:53 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

If have a such website, everybody can add modify complement perfect knowledge base together, write example code, like this convenient.

 Do you have a help file? If you can control OCX or COM DLL, Windows API as a database knowledge base, list all properties, methods, events, constant values, optional parameters, meaning. An example call for each method. Then automatically generated HTM format help file is very convenient, you need a tool like this. Rc6. CWEBVIEW2, there are 70 or so method properties, if you can make a complete example, including all the method calls is perfect, in the code marked: ' 1/AcceleratorKeyPressed/event' 2/AddObject/method' 3/AddScriptToExecuteOnDocumentCreated such a workload is undoubtedly very large, most people use only a small amount of functionality. If do a such general control knowledge base software, Windows API, SQLITE3.0. DLL, VBA Excel. Application and other uses of the knowledge base management, use should be very convenient.

 *[img: Name:  2241RC6.cWebView2谷歌内核浏览器控件&#30340.jpg
Views: 3371
Size:  73.9 KB]*

## #282 tmighty2 — Mar 10th, 2023, 08:08 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

1)
 How would I do this in WebView2, please?
 cefSetting.CefCommandLineArgs.Add("autoplay-policy", "no-user-gesture-required");

 2)
 How could I disable opening new windows?

 3)
 How could I disable that users can right-click the browser to invoke for example "Save as...", "Inspect", etc.?

## #283 jpbro — Mar 10th, 2023, 08:16 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

3)
 How could I disable that users can right-click the browser to invoke for example "Save as...", "Inspect", etc.?

Does setting the WebView2 AreDefaultContextMenusEnabled property to False prevent the menu from appearing?

## #284 jpbro — Mar 10th, 2023, 08:19 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

2)
 How could I disable opening new windows?

Does this work?

Code:

```
Private  Sub WV_NewWindowRequested(ByVal IsUserInitiated As Boolean, IsHandled  As Boolean, ByVal URI As String, NewWindowFeatures As RC6.cCollection)
  IsHandled = True '<- suppress the Browser-Controls default-handler
End Sub
```

See this post: <https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)&p=5597387&viewfull=1#post5597387> https://www.vbforums.com/showthread....=1#post5597387

## #285 Schmidt — Mar 11th, 2023, 12:04 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

1) How would I do this in WebView2, please?
 cefSetting.CefCommandLineArgs.Add("autoplay-policy", "no-user-gesture-required");

Chromium CommandLine-switches can be passed in one of the optional Arguments of the "Bind-call"...
 (and were discussed more than once in this threads history)... *[img: Wink]*
 WV.BindTo(hWnd, , , ,"--autoplay-policy=no-user-gesture-required")

 Here's the Link again, which describes them:
 <https://peter.sh/experiments/chromium-command-line-switches/#autoplay-policy> https://peter.sh/experiments/chromiu...utoplay-policy

 Olaf

## #286 softv — Mar 17th, 2023, 01:30 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Chromium CommandLine-switches can be passed in one of the optional Arguments of the "Bind-call"...
 (and were discussed more than once in this threads history)... *[img: Wink]*
 WV.BindTo(hWnd, , , ,"--autoplay-policy=no-user-gesture-required")

 Here's the Link again, which describes them:
 <https://peter.sh/experiments/chromium-command-line-switches/#autoplay-policy> https://peter.sh/experiments/chromiu...utoplay-policy

 Olaf

// and were discussed more than once in this thread's history //
 Yes indeed, My Dear Olaf. And using this particular tip led me to have moments of great joy once I discovered that **Spell-Checker** (for languages other than the default English) can also be invoked via the Bind-Call. I had to try for quite some time to get it right (I almost decided to write here and ask you) but eventually when I got it working, that was such a joy (**Word suggestions** for misspelt words are also shown!!!). The joy you are giving to so many (few known to you, millions unknown to you) in this world society must be getting registered constantly in the Divine books, I firmly believe. May be it is possible to invoke the translator too! I am yet to explore. If possible and if you know the argument(s) for the same already, kindly let me know. Thanks in advance.

 Well, coming to the Bind-Call (e.g. "--accept-lang=en-US,en,fr" wherein I have included 'fr' also to be considered for spellchecking), I wish to know the following:

 1. For the spellchecker/word-suggestions to function (for any included language), does MS always send the words typed to its servers, to get the required results? Or, is it always done locally after downloading the required language dictionaries once (perhaps updated occasionally too) to user's system? **How "exactly" is this spell-checker/word-suggestions functionality effected by MS in WebView2**? Will be helpful to share this info with users if such info is "concretely" known.

 2. Is it possible to access the word suggestions programmatically and show it in my own separate list?. For instance, when I typed "SpellChecker", the spellchecker engine identifies it as misspelt and when I right click on it, the resultant context menu is as in the picture below. **Is it possible for me to programmatically access these three word suggestions** (Spellchecker, Spellcheckers, Spell-checker) and show them alone in my own list on right-clicking (OR any other way, so that the default right-click menu gets displayed as usual)?

 *[img: Name:  SpellChecker.png
Views: 2957
Size:  4.2 KB]*

 3. I noticed that if I use more than one reference (say WV and WV2) to cWebView2 in my program, then while binding, I needed to "exactly" give the same "--accept-lang" argument in both Bindings. i.e. it had to be as follows:
 --
 If WV.BindTo(picWV.hWnd, TIME_TO_WAIT_m, , , "--accept-lang=en-US,en,fr") = 0
 If WV2.BindTo(picWV2.hWnd, TIME_TO_WAIT_m, , , "--accept-lang=en-US,en,fr") = 0
 --
 Otherwise, the 2nd binding won't happen. I have no problems keeping the same argument in both bind calls, but I feel that perhaps I am doing a mistake and **there possibly is a way by which I can pass different "--accept-lang" argument in each of the WV bind calls**. May be the correct argument is different from "--accept-lang" itself for SpellChecker?

 I searched in this thread right now in each of it's pages and did not find the above matters (with respect to "accept-lang") discussed directly, except one indirect reference to "accept-lang" in the 5th page [xhr.setRequestHeader('Accept-Language', 'en-US,en;q=0.9';]. Hence, I decided to go ahead writing&asking about the above matters via this post. Sure I have done the right thing. In case all of the above have been already discussed (in this thread or elsewhere), then sorry about it. Kindly point me to those messages (here or elsewhere), Olaf. Thank you once again, Olaf, from the bottom of my heart. **God Bless you! God Bless ALL!**

 Kind Regards.

 **Edited**:
 - just corrected "picWV" to "picWV2" in the "WV2.BindTo" line.
 - just a while ago noticed that the picture I inserted is missing. Have inserted it again.

## #287 tmighty2 — Mar 19th, 2023, 06:03 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I would like to access certain elements of a webpage like links, buttons, etc.
 I want to use JavaScript to do that.
 I have uploaded my script here: <https://drive.google.com/file/d/1nfCnyjNdh3IL4639bpHw1O6B_29FTIon/view?usp=share_link> https://drive.google.com/file/d/1nfC...usp=share_link

 First, I need to call grid_ensureHotSpotIds to create these elements.
 I know that this question is relatively specific to JS, but can anybody tell me anyways how I would then access these hotspots from VB6?

## #288 Schmidt — Mar 19th, 2023, 08:11 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

...I need to call grid_ensureHotSpotIds to create these elements.

From what's defined in your js-code, I only see 3 Functions, defined on the (global) window-object:
 window.**_**grid_ensureHotSpotIds
 window.**_**grid_getHotSpots
 window.**_**grid_getFocusedElementHotSpots

 Note the leading underscore of these 3 functions (which you'd have to specify in the function-name-arg of WV.jsRun )

 Here's an example:

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Me.Visible = True 'ensure Visibility before the Bind-call

  Set WV = New_c.WebView2(Me.hWnd)
      WV.AddScriptToExecuteOnDocumentCreated "window._foo = function(){alert('bar')}"
      WV.NavigateToString "<p>dummy-content, nothing to see here...</p>" 'ensure simple "document-content"

  WV.jsRun "window._foo" 'should show an alert('bar')-popup
  WV.jsRun "_foo" '... should show the same (the preceding window. is optional)
End Sub
```

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

...how I would then access these hotspots from VB6?

Just call the functions directly via WV.jsRun (if they are directly callable, due to definition on the Window-Object)...
 Or add little wrapper-functions via WV.AddScriptToExec... (to indirectly call stuff which is deeper hidden behind other objects)

 In either case, please return either "simple types" from these functions (Strings, or Numbers, or Booleans or Nulls) -
 or (in case you want to return js-Arrays or js-Objects) return them this way:
 return JSON.stringify(MyArray_Or_MyObject)

 A function which returns such stringified JSON, can be called this way from VB (auto-decoding the JSON-string already):
 Dim JSON as cCollection
 Set JSON = New_c.JSONDecodeToCollection( WV.jsRun("myfunc", P1, P2) )

 HTH

 Olaf

## #289 softv — Mar 19th, 2023, 10:40 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The following **global EventListener** on the document **works** perfectly for the target element 'abc'.

Code:

```
With New_c.StringBuilder
    .Add "document.addEventListener('change', function(e) {"
    .Add "  e = e||window.event;"
    .Add "  if (e.target.id == 'abc') {"
    .Add "    vbH().RaiseMessageEvent('abcChange', e.target.value)"
    .Add "  };"
    .Add "}, false);"
    WV.AddScriptToExecuteOnDocumentCreated .ToString
End With
```

But, the same EventListener when **applied specifically to the target element** 'abc' (as follows) **does not work** (i.e. the 'abcChange' event is not raised).

Code:

```
WV.AddScriptToExecuteOnDocumentCreated "var abc = document.getElementById('abc'); abc.addEventListener('change', function(event) {vbH().RaiseMessageEvent('abcChange', abc.value);},false);"
```

I think the 'abc' element is not yet loaded when the above script is added. **Am I right?**

 Of course the above line (modified suitably as follows) and included **in the <script> section** of the html file **works** perfectly.

Code:

```
var abc = document.getElementById('abc'); abc.addEventListener('change', function(event) {alert (abc.value + ' eureka');},false);
```

So, to make the above line work through my application itself, I added the following line in the "**WV_DocumentComplete**" event and it **works** but it takes effect after the 2nd 'change' event onwards only. First time when there is change in the 'abc' element, the 'abcChange' event is not raised.

Code:

```
WV.ExecuteScript "var abc = document.getElementById('abc'); abc.addEventListener('change', function(event) {vbH().RaiseMessageEvent('abcChange', abc.value);},false);"
```

So, what is the right way to proceed if I need to add an EventListener to a specific target element alone? **Can somebody kindly educate and help me out? Please**.

 Kind regards.

## #290 softv — Mar 19th, 2023, 12:19 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

In order that the **scripts added to** a WV (via functions like AddScriptToExecuteOnDocumentCreated, etc.) for the **1st page** I load (e.g. God.html) do **not get effected in a 2nd** and different **page** (e.g. Compassion.html) loaded in the same WV, I am right now loading the 2nd page in the manner below. **Is that correct?** If not, what is the correct method?

Code:

```
If WV.BindTo(picWV.hWnd) <> 0 Then
   WV.Navigate App.Path & "\Cosmos\" & "Compassion.html"
End If
```

If I load the 2nd page straightaway using just the following line of code alone, then the scripts added to the 1st page get effected in the 2nd page also if I have elements in the 2nd page with the same 'id's (or classes) as in the 1st page.

Code:

```
WV.Navigate App.Path & "\Cosmos\" & "Compassion.html"
```

Note: I looked for any 'RemoveScript' methods so that by executing them I can remove the scripts already added to a WV so that I can use the abovementioned line of code alone directly. But, I could not find any. I am yet to get time to explore as to what RemoveObject method does.

 God Bless you, olaf! **God Bless ALL!**

 Kind regards.

## #291 Schmidt — Mar 19th, 2023, 12:24 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>
 So, what is the right way to proceed if I need to add an EventListener to a specific target element alone?

Why not use the global Listener, when it works nicely already?

 I'd make it even more generic, so that it works for any Widget which raises a change-event:
 (this way you can filter on the VB-side)...

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Me.Visible = True 'ensure Visibility before the Bind-call

  Set WV = New_c.WebView2(Me.hWnd)

  With New_c.StringBuilder
      .AddNL "document.addEventListener('change', (e)=>{"
      .AddNL "    vbH().RaiseMessageEvent(e.target.id+'_change', e.target.value)"
      .AddNL "})"
      WV.AddScriptToExecuteOnDocumentCreated .ToString
  End With

  WV.NavigateToString "<input id='abc' type='text'><br><br><textarea id='xyz'>"
End Sub

Private Sub WV_JSMessage(ByVal sMsg As String, ByVal sMsgContent As String, oJSONContent As RC6.cCollection)
  Select Case LCase$(sMsg)
    Case "abc_change": Debug.Print "abc was changed to:"; sMsgContent
    Case "xyz_change": Debug.Print "xyz was changed to:"; sMsgContent
  End Select
End Sub
```

Olaf

## #292 softv — Mar 19th, 2023, 01:02 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Why not use the global Listener, when it works nicely already?

 I'd make it even more generic, so that it works for any Widget which raises a change-event:
 (this way you can filter on the VB-side)...
 ... .. . ... .. .
 Olaf

// Why not use the global Listener, when it works nicely already? //
 **Yes, I definitely shall, olaf**. I needed such a confirmation from you. Because, otherwise, I was somewhat concerned whether listening to events at the document level would increase the load on the 'events manager'. Now that you have given your assured go-ahead with a **beautiful tip of a generic method** itself, I am confident now that I need not have any concerns as aforesaid. **Thanks a TON, as always**.

 Kind regards.

## #293 tmighty2 — Mar 19th, 2023, 01:58 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I think my question was misleading. I would like to ask how I could then interact with the hotspots that I assigned.
 Does your framework perhaps offer a built-in method to iterate over them? I am not yet sure how I should access them, for example to get all the links in a webpage or to assign a different color to a hotspot area.

## #294 softv — Mar 22nd, 2023, 12:03 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **softv** *[img: View Post]*

// and were discussed more than once in this thread's history //
 ... .. .
 Well, coming to the Bind-Call (e.g. **"--accept-lang=en-US,en,fr"** wherein I have included 'fr' also to be considered for spellchecking), I wish to know the following:

 1. For the spellchecker/word-suggestions to function (for any included language), does MS always send the words typed to its servers, to get the required results? Or, is it always done locally after downloading the required language dictionaries once (perhaps updated occasionally too) to user's system? **How "exactly" is this spell-checker/word-suggestions functionality effected by MS in WebView2**? Will be helpful to share this info with users if such info is "concretely" known.

 2. Is it possible to access the word suggestions programmatically and show it in my own separate list?. For instance, when I typed "SpellChecker", the spellchecker engine identifies it as misspelt and when I right click on it, the resultant context menu is as in the picture below. **Is it possible for me to programmatically access these three word suggestions** (Spellchecker, Spellcheckers, Spell-checker) and show them alone in my own list on right-clicking (OR any other way, so that the default right-click menu gets displayed as usual)?

 *[img: Name:  SpellChecker.png
Views: 2957
Size:  4.2 KB]*

 3. I noticed that if I use more than one reference (say WV and WV2) to cWebView2 in my program, then while binding, I needed to "exactly" give the same "--accept-lang" argument in both Bindings. i.e. it had to be as follows:
 --
 If WV.BindTo(picWV.hWnd, TIME_TO_WAIT_m, , , "--accept-lang=en-US,en,fr") = 0
 If WV2.BindTo(picWV2.hWnd, TIME_TO_WAIT_m, , , "--accept-lang=en-US,en,fr") = 0
 --
 Otherwise, the 2nd binding won't happen. I have no problems keeping the same argument in both bind calls, but I feel that perhaps I am doing a mistake and **there possibly is a way by which I can pass different "--accept-lang" argument in each of the WV bind calls**. May be the correct argument is different from "--accept-lang" itself for SpellChecker?
 ... .. .

 Kind Regards.

With respect to my above post, Since Olaf may not be finding time always, can somebody else (dear JpBro, dear Wwolf, ..., .., .), if at all their time and situation permits, **kindly help me out providing answers to my above 3 questions.**

 There is also a sub-question in my 3rd question above. Let me pose it as a separate question hereunder so that it is easier for anyone to answer.
 --
 I have used **"--accept-lang=en-US,en,fr"** in the 'BindTo' call to enable and make use of the SpellChecker/WordSuggestions facility of Edge-WebView2 for languages other than the default English. **Is that enough?** **Is that exactly the right way?** If so, has anybody found time to use it extensively? And, if so, is the Spellchecking/WordSuggestions facility working well under all situations?
 --

 Sorry that I have not found time yet to explore the Spellchecking/WordSuggestions facility extensively, myself. Also, first I need to get confirmed from Olaf or others that **"--accept-lang=en-US,en,fr" **is enough and its the right way to enable and make use of the Spellchecking/WordSuggestions facility.

 **Thanks in advance.**

 Kind Regards.

## #295 Schmidt — Mar 22nd, 2023, 02:15 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **softv** *[img: View Post]*

...kindly help me out providing answers to my above 3 questions.

...as for 1)
 MS-WebView2 is still "just a relatively thin wrapper" around the chromium -opensource-project (mainly managed by google) - and therefore all the main-functionalities (as e.g. spellchecking and its context-menus) are in all likelihood "chromium-implementations".

 There might be an exception here, regarding the necessary dictionaries (compared to chromium-CEF-binaries on e.g. Linux).
 I'm not sure whether MS is redirecting to the "local OS-dictionaries, which come with the Win-language-packs" -
 this might be the case, to keep the "Edge-Evergreen-package" relatively small... though I'm not sure about that...

 "Constant online-access" (to do word-lookups) does not seem to happen in my tests...
 (have temporarily disabled any Inet-connections here - and get proper spell-suggestions for "de" and "en" nevertheless.

 ...as for 2)
 There's currently no "context-menu-parsing" exposed in the RC6.cWebView2-class -
 (although interfaces for that exist in the MS-typelib).

 There's also no way, to retrieve the contextmenu-contents at "js-level".

 What is supported currently is, to entirely suppress any context-menu-popups (AreContextmenusEnabled) -
 and then show your own, by using the appropriate ContextMenu-EventHandler (according to the transported x/y-pos).

 ...as for 3)
 Was astonished by that myself (that the commandline-args have to remain constant for a Process-session) -
 but have not tried to find a reason or a workaround, because I can live with the behaviour...
 Perhaps changing the "UserDataFolder" in the Bind-call for the second instance could help, but haven't tested that either.

 Generally, the browser-internal spellchecking functionality is not well-exposed
 (only very few things can be influenced via javascript),
 and that's the reason why there's such an amount of js-libs which address this
 (usually in conjunction with some <textarea>bindings).

 It's up to you to use some of these libs (or the "big-guns" like CKEditor or TinyMCE, which come with their own spellcheckers as well).

 Olaf

## #296 Schmidt — Mar 22nd, 2023, 02:54 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

I would like to ask how I could then interact with the hotspots that I assigned.
 Does your framework perhaps offer a built-in method to iterate over them?

Why not iterate over them from the inside of a little js-function?
 You can pass String-Parameters into such functions just fine from the VB-Side...
 (like e.g. QuerySelector-StringPatterns, or just "plain IDs", along with e.g. Color-Strings).

 In short, complex DOM-operations remain "in generic, clever parametrized js-helpers" -
 and then these "nice trigger-parameters" could (should) be passed from the VB-side.

 Olaf

## #297 softv — Mar 22nd, 2023, 06:34 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

...as for 1)
 MS-WebView2 is still "just a relatively thin wrapper" around the chromium -opensource-project (mainly managed by google) - and therefore all the main-functionalities (as e.g. spellchecking and its context-menus) are in all likelihood "chromium-implementations".

 There might be an exception here, regarding the necessary dictionaries (compared to chromium-CEF-binaries on e.g. Linux).
 I'm not sure whether MS is redirecting to the "local OS-dictionaries, which come with the Win-language-packs" -
 this might be the case, to keep the "Edge-Evergreen-package" relatively small... though I'm not sure about that...

 "Constant online-access" (to do word-lookups) does not seem to happen in my tests...
 (have temporarily disabled any Inet-connections here - and get proper spell-suggestions for "de" and "en" nevertheless.

 ...as for 2)
 There's currently no "context-menu-parsing" exposed in the RC6.cWebView2-class -
 (although interfaces for that exist in the MS-typelib).

 There's also no way, to retrieve the contextmenu-contents at "js-level".

 What is supported currently is, to entirely suppress any context-menu-popups (AreContextmenusEnabled) -
 and then show your own, by using the appropriate ContextMenu-EventHandler (according to the transported x/y-pos).

 ...as for 3)
 Was astonished by that myself (that the commandline-args have to remain constant for a Process-session) -
 but have not tried to find a reason or a workaround, because I can live with the behaviour...
 Perhaps changing the "UserDataFolder" in the Bind-call for the second instance could help, but haven't tested that either.

 Generally, the browser-internal spellchecking functionality is not well-exposed
 (only very few things can be influenced via javascript),
 and that's the reason why there's such an amount of js-libs which address this
 (usually in conjunction with some <textarea>bindings).

 It's up to you to use some of these libs (or the "big-guns" like CKEditor or TinyMCE, which come with their own spellcheckers as well).

 Olaf

**Thanks a TON**, as always, dear Olaf.

 And, before anything else, I just presume from your detailed answers that "--accept-lang" is indeed the right way, as of now, to add SPWS (short for SpellChecker/WordSuggestions) facility via WebView2 in our own applications and that is enough too (though SPWS via "--accept-lang" does not seem to work for some languages; please see details in my "observations" below). So, I can very much go ahead using "--accept-lang" in my applications to enable and make use of SPWS (for whatever languages it works for now). **Right?** **Kindly confirm.**

 Well, I took time to explore more on this SPWS (short for SpellChecker/WordSuggestions) facility since the past 5 hours or so. And, noted down the following observations. As I was doing that, when I checked in between, I saw your answers to my 3 questions, that too in such great detail. Thanks a TONNNN, again, dear Olaf. Well, a few(not all) portions of my following observations may have been already responded to by you in your recent answers but if I start rephrasing my following observations, it may turn out to be a difficult/confusing process for me since already I have spent quite some time phrasing and rephrasing them in an understandable manner. So, I am sharing my observations just as I originally started noting them down itself and not modifying them based on your recent answers. Kindly bear with me for the same. Thanks in advance for your understanding in this regard.

 My observations:
 --
 1. When using "--accept-lang" to enable and use SPWS, some languages work and others don't. For instance, "fr,de,es,ta,hi,ml,te" work but NOT "ar,sa,kn,mr,gu,pa,as,bn". To confirm that all the language codes I am specifying are correct only, I referred to this page - <https://www.biswajeetsamal.com/blog/web-browser-language-identification-codes> https://www.biswajeetsamal.com/blog/...fication-codes. It came first on google search for "browser language codes" and I just used it. That's all.

 2. As an example for the abovementioned observation, if "--accept-lang=ar-AE,es,kn,en-US,en,fr-FR,de,as,bn" is specified in the BindTo call, the "Check spelling" context menu shows "Spanish, English (United States), English, French (France), German)" alone. i.e. only "es, en-US, en, fr-FR, de" are considered for spell-checking. "ar,kn,as,bn" are not considered. Please see screenshot at the end of this message (I opted to place the screenshot at the end so that the flow of this long message is not disturbed while reading).

 3. I just tried providing a list of non-working languages alone to "-accept-lang". That did not work. I tried placing the non-working languages in a different order in the list, just in case. That did not help either.

 4. For some sets of languages specified, I noticed that if a language which does NOT work is found in the list, even a language for which SPWS works is not shown in the "Check spelling" context menu.
 For example, if "--accept-lang=en,fr,ta,hi,te,ml,mr" is specified in the BindTo call, the "Check spelling" context menu shows "English, French, Tamil, Hindi" (en,fr,ta,hi) alone. "te,ml" are not shown. This is because of the 'mr'(Marathi) after 'ml'(Malayalam). In other words, if 'mr' is omitted and if "--accept-lang=en,fr,ta,hi,te,ml" is specified in the BindTo call, the "Check spelling" context menu does show all of "English, French, Tamil, Hindi, Telugu, Malayalam".

 5. In my Microsoft Edge browser (updated to the latest - 111.0.1661.51), under 'Settings->Languages->User writing assistance', it shows two modes of SPWS - 'Microsoft Editor' mode and 'Basic' mode. It is mentioned therein that if the former is used, "**enhanced**" SPWS is done by MS and the typed words are sent to MS for processing. If the latter is used, typed words are processed locally. What mode possibly WebView2 could be using?

 6. Since the SPWS seem to work when the internet is OFF also, only the 'Basic' mode is used by WebView2? If so and if 'Microsoft Editor' mode has to be used for "enhanced" SPWS, what should be done (**if at all it can be done** with the current WebView2 itself)? I just checked the SPWS with a limited set of words only for a few languages both with internet ON and OFF. Need to check with 5 to 6 sets of BIG chunk of words for the same few languages at some later point of time, to see whether the same misspelt words are shown underlined both when internet is ON and OFF.

 7. For some languages, MS Edge itself says that SPWS is not yet supported. For e.g. for "Bangla(India)".

 8. I noticed that sa(Sanskrit) was not found in the original list (of languages provided by MS Edge to add for SPWS) itself.

 9. <https://github.com/nwjs/nw.js/issues/4954> https://github.com/nwjs/nw.js/issues/4954 - in this page, this comment (<https://github.com/nwjs/nw.js/issues/4954#issuecomment-607348899> https://github.com/nwjs/nw.js/issues...ment-607348899) suggests a summary of 3 very simple things to do to add SPWS facility via package.json and Javascript. This thread was in 2017 though. So, I don't know whether the same is applicable via JavaScript to WebView2 too. **If it is applicable** and if it can serve as an additional way (apart from "--accept-lang") to enable and use SPWS for different languages, then, when your time permits, kindly elucidate more on the same, because I am not familiar with package.json, etc. May be SPWS for the currently non-working languages (ar,kn,gu,pa,etc.) will work through the aforementioned JS way?

 10. Even if I remove all the languages added through my MS Edge, I was still able to add SPWS facility via WebView2 in my application. So, MS language settings do not seem to have any connection with the language settings added via BindTo call in my application. Only thing is that WebView2 does not seem to support SPWS for all languages, for whatever reasons. Or perhaps I am doing some mistake. If I am doing some mistake, then, **kindly point out** the same so that I am able to enable and use SPWS facility for the currently non-working languages (ar, kn, gu, pa, etc.) too via "--accept-lang".

 11. Thanks for your helpful suggestion on js-libs for SPWS. Actually, I also noticed (just today, during my above exploration) that some js libs are available for spellchecking. I did not explore what they are though. If they are of MIT License supporting many languages, that would be great. Otherwise, for the kind of freeware-only work I do, whatever languages currently supported by WebView2 are itself such an "extremely exciting" thing. Really. **Thanks** for your suggestions, once again.
 --

 All said and done, as requested at the start of this message, **kindly confirm**, dear Olaf, that I can very well go ahead using "--accept-lang" in my applications to enable and make use of SPWS (for whatever languages it works for now).

 Once again, TONs of Thanks from the bottom of my heart for such prompt and detailed replies to each and every doubt of each and every user of RC6 in various threads related to RC6. And, those invaluable tips and cute compact code snippets now and then, to help us along with. God Bless you, dear Olaf. **God Bless All**.

 Kindest Regards.

 Screenshot (for point no.2 above):
 *[img: Name:  SpellChecker-2.png
Views: 2637
Size:  7.9 KB]*

## #298 tmighty2 — Mar 22nd, 2023, 07:25 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I make a software for handicapped people.
 I need to get the links and clickable button elements of a webpage in order to show them on extra big buttons in my software.
 Instead of clicking in the website, handicapped users can press the big buttons in my software.
 At a first click on a button, the element in the webpage will be highlighted with a red outline, and its text will be spoken by a computer voice.
 At a 2nd click on the button in my application, the link of them element will be navigated to.

 That is why I need to be able to iterate through these elements from outside.

 For example, when the user presses button #2, I need to highlight hotspot #2 in the webpage.

 Can you tell me what you mean by "complex DOM-operations remain "in generic, clever parametrized js-helpers"?
 If possible, could you show me how to iterate over these hotspots from within VB6 using JS?

 I don't know how I could get back an array of hotspots from the browser by calling the JS function. I hope I could explain it well.

## #299 softv — Mar 22nd, 2023, 12:59 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Note: This is a follow-up message to my post No. 297 ("11-point observations" message).

 First thing first, Olaf. SPWS **works for** **all the languages** with "--accept-lang" argument. *[img: Smilie]*

 What happened was that after a while after sending my "11-point observations" message to you, it occurred in my mind that I was all the while concentrating on adding languages to MS Edge rather than checking the Language packs installed in my system. At that point of time, **your wise statement** (I'm not sure whether MS is redirecting to the "local OS-dictionaries, which come with the Win-language-packs") also came to my mind simultaneously.

 So, I checked the language packs installed and found out that language packs for only a few languages (out of the many languages I was checking out) were installed in my Win10 system. So, I started installing the missing language packs but for whatever reasons, none of the language packs got installed, however many times I tried. That is when it occurred to me that I can easily check out the compiled exe in one of my other Win10 systems which I use only now and then (but keep its OS upto-date with the latest OS updates whenever I use it).

 So, I switched on my other system and checked out the compiled exe. And, all languages worked straightaway! *[img: Smilie]*. I checked the language packs installed in my other system and all the required language packs had been installed in it. So, as of now, SPWS works for all the languages with "--accept-lang" argument. The thing is that, as guessed by you very correctly (as it is so usually *[img: Smilie]*), the language packs for all the languages for which SPWS is needed should be installed priorly in the system. Thanks a TON ton ton, my dear Olaf.

 Well, in the complied exe I had "--accept-lang=ta,hi,sa,te,kn,ml,mr,kok,gu,bn,pa,or,as,ar,es,en-US,en,fr-FR,de" and all of them (**19 languages**) showed up in the "Check spelling" context menu.

 One more thing. In my eagerness to check out the working of SPWS in my other system, I forgot the fact that I was checking out the compiled exe in one other system (in which RC6 is not installed/registered) for the very first time. So, when the starting screen of the compiled exe appeared, only after a second I realised that the compiled exe started "registration-free". Hats off to you, Olaf. This is the first time I am running a "compiled exe with DLL dependencies" and yet seeing it opening without me having to register the DLLs myself. Your DirectCOM.dll does it automatically. I did use your **"modRCRegFree.bas" drop-in **in my application, just exactly as you had suggested. *[img: Smilie]*. Thanks once again.

 All said and done, as requested in my earlier post (11-point message), when possible, **kindly give your confirmation** that I can very well go ahead using the "--accept-lang" approach in my applications to enable and make use of SPWS and no issues are foreseen in this approach/method. Thanks in advance. Any helpful/useful/important tips that you wish to give me when I go forward using this "--accept-lang" approach for SPWS, kindly let me know please. Looking forward to the same too. God Bless you, olaf. **God Bless ALL**.

 Kind Regards.

## #300 softv — Mar 22nd, 2023, 01:38 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

By the way, Olaf, thanks (as usual) A TONNNN for your most wonderful tip on **VSCODE editor** (in post no. 277).

 What a gem this little cute editor is! I was thinking that VSCODE would be yet another bulky IDE (like Visual Studio) where I would be taking several days to learn (and also remember) the IDE's various menus/options/features itself. But it was not at all so! It was such a pleasant surprise and joy to see VSCODE being so light-weight (yet features-packed!). I never knew that such a tailor-made cute little app can exist for html5/css/js coding. It has all the features (and more *[img: Smilie]*) one would expect in an html/css/js editor. Simply Fantastic. Happy that I got to see this helpful tip of yours early. **Your superb tip** has saved me so many hours thereafter and increased my productivity a lot. You are so special for our VBForums. Thanks once again. *[img: Smilie]*.

 Kindest Regards.

## #301 softv — Mar 24th, 2023, 05:19 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Dear Olaf,

Code:

```
WV.AddSCriptToExecuteOnDocumentCreated New_c.FSO.ReadTextContent("MyHelpers.js", False, CP_UTF8)
```

With reference to your above one line of code in post no. 277, I have the following query.
 --
 When trying to add a .js file (say 'simple.js') in the above manner, I found that it was working only if there are no jQuery calls in my 'simple.js' file. It took me some considerable amount of time to find out that jQuery calls were the problem. Am I doing a mistake? Or, is it that I indeed cannot use **jQuery calls** in the js strings added via WV.AddSCriptToExecuteOnDocumentCreated? If I can use jquery, then how to achieve that? I do have proper references to jquery.js in my 'simple.html' file. 'Simple.html' does work correctly if I include 'simple.js' directly in it and I "WV.Navigate" to 'simple.html' directly in my VB application.
 --

 Note: I did take time to search each of the pages in this thread for 'jQuery'. But, I could not find any post discussing about it. I thought may be it is written in some post by dear Olaf that only 'Pure JS' can be used. So, right now I searched on "Pure JS" in this thread but could not find any post mentioning the same. Well, may be it is mentioned in a different way by dear Olaf. I don't know.

 In the above context, I will be truly grateful if someone can let me know what is the best way to find out whether a particular matter has been already discussed in this thread (in this thread "alone"). At present, I am clicking each numbered page link at the top of this thread and searching on a particular term (say 'jQuery'). Is there a **faster **way of** searching** so that if I search on 'jQuery', I immediately get to see all the pages (in this thread alone) in which it appears or 'no results' if it is not present in any of the pages (in this thread). That will help me a lot. Really. I did try the 'Advanced search but it does not seem to zero-in on the specific post(s). For e.g. for a 'Keyword:' of FSO.ReadTextContent and 'User Name:' of Schmidt, the results do not show me the exact post (277) in which the Keyword has been used. May be I am searching wrongly, I don't know. Sometimes, I do miss the obvious and focus elsewhere. Probably that is the case here also. So, kindly help me out.

 I take this opportunity to thank you once again, Olaf, for helping us all out, to the best extent you can, always, even amidst your hectic schedules. I personally know how arduous it is (to give such help, regularly) but at the same time how joyful it is too (ultimately) since it results in immense benefit to the world society. God Bless you. God Bless ALL.

 Kind Regards.

## #302 Schmidt — Mar 24th, 2023, 07:43 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **softv** *[img: View Post]*

...When trying to add a .js file (say 'simple.js') in the above manner,
 I found that it was working only if there are no jQuery calls in my 'simple.js' file.

With the now available browser-built-in selector-functions, there's rarely a need for jQuery anymore...

 That said - jQuery is an external js-lib - and therefore needs to be (pre)loaded as well -
 ideally before you add your own js-scripts.

 You can ensure that either in the HTML (via the usual <script>-tags) -
 or use AddScriptOnDocumentCreate as well for jQuery...

 Below is some Demo-Code:
 (which downloads jQuery dynamically "as String" via a http51-helperfunction - but once downloaded via wgetJQ(),
 you can of course put the retrieved string into your own local js-file - and load it from there in future-runs)...

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Me.Visible = True

  Set WV = New_c.WebView2(hWnd)
      WV.AddScriptToExecuteOnDocumentCreated wgetJQ("https://code.jquery.com/jquery-3.6.4.js")
      WV.AddScriptToExecuteOnDocumentCreated "function setColor_SelectorBased(sel,clr){" & _
                                             "   $(sel).each(function(){$(this).css('background-color',clr)})" & _
                                             "}"

  WV.NavigateToString "<p>first Para</p><p>second Para</p><p>last Para</p>"

  WV.jsRun "setColor_SelectorBased", "p:first, p:last", "magenta"
End Sub

Public Function wgetJQ(URL) As String
  With CreateObject("WinHttp.WinHttpRequest.5.1")
      .open "GET", URL: .send: wgetJQ = .responseText

      'hot-fixes for two edge/chromium-related bugs in the current jQuery-release
      wgetJQ = Replace(wgetJQ, "el = document", "el = window.document")
      wgetJQ = Replace(wgetJQ, "documentElement.getRootNode", "documentElement && documentElement.getRootNode")
  End With
End Function
```

If all works well, the first and the last of the 3 <p>-Tags should be colored magenta.

 Edit: Seems, that the WebView2-method: AddScriptToExecuteOnDocumentCreated -
 does not live up to its "second symbolname-part" anymore...
 (hope, they are fixing this soon)...

 In the interim, you can enclose all such added scripts with your **own js-event-listener**:
 (example is based on the working one above, but now not even the "jQuery-hotfixes" are needed anymore,
 when the script-adding is **delayed **until the document "is sitting properly in the DOM")

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Me.Visible = True

  Set WV = New_c.WebView2(hWnd)
  With New_c.StringBuilder
      .AddNL wget("https://code.jquery.com/jquery-3.6.4.js")

      .AddNL "$('p:first, p:last').each(function(){$(this).css('background-color','magenta')})"

      WV.AddScriptToExecuteOnDocumentCreated "window.addEventListener('load',(e)=>{" & .ToString & "})"
  End With

  WV.NavigateToString "<p>first Para</p><p>second Para</p><p>last Para</p>"
End Sub

Public Function wget(URL) As String 'generic text-retrieval "by URL"
  With CreateObject("WinHttp.WinHttpRequest.5.1")
      .open "GET", URL: .send: wget = .responseText
  End With
End Function
```

Olaf

## #303 tmighty2 — Mar 24th, 2023, 07:51 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Why not iterate over them from the inside of a little js-function?
 You can pass String-Parameters into such functions just fine from the VB-Side...
 (like e.g. QuerySelector-StringPatterns, or just "plain IDs", along with e.g. Color-Strings).

 In short, complex DOM-operations remain "in generic, clever parametrized js-helpers" -
 and then these "nice trigger-parameters" could (should) be passed from the VB-side.

 Olaf

Sorry to ask again, but can you please show me an example?

 My problem is that I don't know how to access the hotspots that the function returns.
 I have tried this:

 Dim hotspots
 hotspots = WV.jsRun("window._grid_ensureHotSpotIds")

 Debug.Print UBound(hotspots) 'here it errors out saying "Type mismatch" (obviously it's not an array. It's empty)

## #304 softv — Mar 24th, 2023, 10:52 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

With the now available browser-built-in selector-functions, there's rarely a need for jQuery anymore...

 ... .. .

 Edit: Seems, that the WebView2-method: AddScriptToExecuteOnDocumentCreated -
 does not live up to its "second symbolname-part" anymore...
 (hope, they are fixing this soon)...

 In the interim, you can enclose all such added scripts with your **own js-event-listener**:
 (example is based on the working one above, but now not even the "jQuery-hotfixes" are needed anymore,
 when the script-adding is **delayed **until the document "is sitting properly in the DOM")

Code:

```
... .. .
```

Olaf

Dear Olaf,

 Thanks a TON for both the initial demo-code and its Edit later on.

 The starting info mentioned by you in your "Edit" portion was exactly the cause of my problem (reported earlier) and not jQuery.

 Actually, after posting my earlier message, as I was exploring further, I realised that even some "pure js" statements were not working (much like the jQuery statements) if they were lying globally (instead of inside some functions) in my ".js" file. So, what I did finally was to place all of those non-working lines of code inside the following **DOMContentLoaded** "listener" function

Code:

```
Document.addEventListener('DOMContentLoaded', function () {
 ... .. .
}
```

Then, everything worked "perfectly" with absolutely no hitch. That made me come to the realisation that the DOM elements were not visible to the globally lying lines of code if my whole js file is added as a string via 'AddScriptToExecuteOnDocumentCreated'. Based on that realisation, the conclusion to which I reached was that my understanding of what "AddScriptToExecuteOnDocumentCreated" would do was incorrect.

 I came back here to share my conclusion and asking you whether my understanding is correct. That is when I saw your Edit wherein you have remarked that the 2nd part of "AddScriptToExecuteOnDocumentCreated" is a sort of bug itself. Thanks Olaf.

 And, it was very nice to see the following 'yet another' cute compact one-liner *[img: Smilie]* from you to overcome the seeming bug in "AddScriptToExecuteOnDocumentCreated".

Code:

```
WV.AddScriptToExecuteOnDocumentCreated "window.addEventListener('load',(e)=>{" & .ToString & "})"
```

So, what I did (as a very quick test) was to check out your above one-liner straightaway in my VB application [of course after replacing the ".ToString" in your one-liner with "New_c.FSO.ReadTextContent("MyFullFile.js", False, CP_UTF8)"]. But, it looked like that the above resulted in the existing function [Document.addEventListener('DOMContentLoaded', function ()] in my ".js" file to not to take effect. I remember reading somewhere in the net today (and some days back too) that if we decide to use more than one OnLoad listener, then these listerners should be added to a list of listeners. OR else, sometimes, one of the listeners would replace the other listeners. Perhaps that is what is happening when I use your one-liner. I don't know that for sure now. Because, as of now, I just did a very quick test only. Thanks once again for your cute one-liner. I just loved it. When time permits, I will try to learn what to do in my ".js", in order to use your one-liner.

 By the way, when you say "With the now available browser-built-in selector-functions, ... .. .", I think you mean the querySelector and querySelectorAll functions. If so, then, yes, after reading your posts mentioning about them, I do have started using them and are very convenient. Will try to slowly move over to querySelectors completely. Thanks a TON.

 Kind Regards.

## #305 Schmidt — Mar 24th, 2023, 12:17 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **softv** *[img: View Post]*

...it looked like that the above resulted in the existing function
 [Document.addEventListener('DOMContentLoaded', function ()] in my ".js" file
 to not to take effect. ...

If it is enclosed by "my" outer "onload"-callback-snippet, then "your" (inner) event-listener-installation comes "too late" ...
 since you prepare your listening, when the document-load-event has already happened (was already catched by me) -
 and therefore installing the Listener "at that point in time" will never trigger the code in "your" inner callback-function-block...

 Easy solution would be, to remove your own listener-installation from your js-source-file
 (also in preparation for a potentially upcoming bugfix by MS for the AddScriptToExecuteOnDocumentCreated-call).

 Olaf

## #306 Schmidt — Mar 24th, 2023, 12:27 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

Sorry to ask again, but can you please show me an example?

 My problem is that I don't know how to access the hotspots that the function returns.
 I have tried this:

 Dim hotspots
 hotspots = WV.jsRun("window._grid_ensureHotSpotIds")

 Debug.Print UBound(hotspots) 'here it errors out saying "Type mismatch" (obviously it's not an array. It's empty)

You will have to check in your js-source, whether the window._grid_ensureHotSpotIds function,
 returns something at all...

 And if it does, then gather these results in your own array -
 and then return this array as a serialized JSON-string via:
 return JSON.stringify(myHotspotArray)

 At the vb6-side you can then easily convert that returned JSON-string into a cCollection:
 Dim oHotSpots As cCollection
 Set oHotSpots = New_c.JSONDecodeToCollection(WV.jsRun("window._grid_ensureHotSpotIds"))

 HTH

 Olaf

## #307 softv — Mar 25th, 2023, 05:46 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

... .. .
 Easy solution would be, to remove your own listener-installation from your js-source-file
 (also in preparation for a potentially upcoming bugfix by MS for the AddScriptToExecuteOnDocumentCreated-call).
 Olaf

What care and concern from you in your support!!! I am quite enjoying it, working with RC6 and also receiving your wonderful loving support. Have missed a LOT of both, I am sure, by not getting time to start using 'RC' very many years back itself, though I always knew inside my heart that 'RC' was such a treasure of a contribution!

 Well, the very first thing I tried was that only - removing my own 'DOMContentLoaded' listener line. But, it did not work. That quite surprised me, like anything! Because, I was almost certain that that would be enough. So, I started examining my code carefully and found out that I had left a piece of helper-function code still in my VB app itself. I commented out that line and added that function also in my 'MyFullFile.js' and **everything worked like a charm**!

 Your tip in post 277 that all helper functions shall be called from a single .js file made me think why not call my whole .js file itself from inside VB. Now that it is possible, it is (and will continue to be) of great help to me, going forward, I believe. Once again thanks for your OnLoad listener one-liner! Superb!

 Ever remaining in awe of both your work and your heart.

 Kindest Regards.

## #308 tmighty2 — Mar 25th, 2023, 02:11 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

My problem is that stringify does not work. hotspots.count is 196, but the line "alert('after stringify');" is not fired in the code below:

Code:

```
HTMLElement.prototype._grid_ensureHotSpotIds = function(selector)
{
  let hotSpots = [];
  if (selector != null)
  {
    var splitSelector = selector.split(',');
    //Some websites need a selector that's too complex to make just in one step. So we need to split the selector on commas and do querySelectors that will chain the results.
    splitSelector.forEach(singleSelector =>
    {
      if (hotSpots != null)
      {
        let hotSpotsFound = this.querySelectorAll(singleSelector);
        if (hotSpotsFound != null)
        {
          hotSpots = hotSpots.concat(Array.from(hotSpotsFound));
        }
        else
        {
          //alert('hotspots is null in splitselector');
        }
      }
    });
  }
  else
  {
    hotSpots = this.querySelectorAll(hotSpotSelector);
  }
  //alert("Found " + hotSpots.length + " hotspots."); // Alert showing the number of hotspots found

  for (var hotSpot of hotSpots)
  {
    hotSpot._grid_ensureId();
  }
  alert('before stringify');
  let hotSpotsString = JSON.stringify(hotSpots); // Convert the hotspots array to a JSON string
  alert('after stringify');
  alert("Hotspots: " + hotSpotsString); // Alert showing the hotspots as a string

  return hotSpotsString; // Return the hotspots array as a JSON string
};
```

I guess it could be due to members of such a hotspot.
 I have uploaded the prototype of "element" which I try to stringify:

 <https://drive.google.com/file/d/1XPd-5MTLVtz6y8Jc3_w5U16v1LolZEnS/view?usp=share_link> https://drive.google.com/file/d/1XPd...usp=share_link

 Edit: I know now that in C# the hotspots are handled like this:

 C# code:

Code:

```
protected virtual Task<HtmlElementReference> GetStartElementAsync(HtmlElementReference startElement, CursorDirection direction)
		{
			return Task.FromResult<HtmlElementReference>(startElement);
		}

		protected async Task<HtmlElementReference> HandleFindElementResultAsync(ElementCursorBase.FindElementResult result, IWebBrowserFrame currentFrame, string getNodesFunc, int? nodeType, CursorDirection direction)
		{
			HtmlElementReference htmlElementReference;
			HtmlElementReference elementReferenceAsync;
			CursorDirection cursorDirection;
			string str;
			if (result != null)
			{
				switch (result.ResultType)
				{
					case ElementCursorBase.FindElementResultType.Element:
					{
						htmlElementReference = new HtmlElementReference(currentFrame.Identifier, result.Element);
						return htmlElementReference;
					}
					case ElementCursorBase.FindElementResultType.IterateUpIntoParent:
					{
						HtmlElementReference elementReferenceAsync1 = await WebBrowserExtensionMethods.GetElementReferenceAsync(currentFrame.Parent, String.Format("document.querySelector('[data-grid-frame-identifier=\"{0}\"]')._grid_ensureId()", currentFrame.Identifier));
						cursorDirection = direction;
						if (cursorDirection == CursorDirection.Backward || cursorDirection == CursorDirection.BackwardToHeading)
						{
							getNodesFunc = "_grid_getPreviousNodes";
						}
						htmlElementReference = await this.FindElementAsync(elementReferenceAsync1, getNodesFunc, nodeType, false, direction);
						return htmlElementReference;
					}
					case ElementCursorBase.FindElementResultType.IterateDownIntoFrame:
					{
						using (IWebBrowserFrame frame = this._webBrowser.GetFrame(Convert.ToUInt64(result.Element)))
						{
							if (frame != null)
							{
								cursorDirection = direction;
								switch (cursorDirection)
								{
									case CursorDirection.Forwards:
									case CursorDirection.ForwardsToHeading:
									{
										str = "document.body._grid_ensureId()";
										break;
									}
									case CursorDirection.Backward:
									case CursorDirection.BackwardToHeading:
									{
										str = "document.body._grid_getLastElement()._grid_ensureId()";
										getNodesFunc = "_grid_getSelfAndPreviousNodes";
										break;
									}
									default:
									{
										throw new InvalidOperationException();
									}
								}
								elementReferenceAsync = await WebBrowserExtensionMethods.GetElementReferenceAsync(frame, str);
							}
							else
							{
								htmlElementReference = null;
								return htmlElementReference;
							}
						}
						frame = null;
						htmlElementReference = await this.FindElementAsync(elementReferenceAsync, getNodesFunc, nodeType, true, direction);
						return htmlElementReference;
					}
				}
				throw new InvalidOperationException();
			}
			else
			{
				htmlElementReference = null;
			}
			return htmlElementReference;
			throw new InvalidOperationException();
		}
```

Would there be any way to do the same with your browser?

## #309 tmighty2 — Mar 27th, 2023, 06:25 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

To be more specific: Can VB6-WebView2 handle async tasks like shown in the C# code?

## #310 tmighty2 — Mar 27th, 2023, 08:31 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

ps: cairo_sqlite.dll is missing a version number.

## #311 tmighty2 — Mar 27th, 2023, 09:19 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I have tried to disable the context command, but I think there are no commandline switches available for these settings:

 webView.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;
 webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;

 How could I do this with VB6-WebView2?

## #312 Schmidt — Mar 27th, 2023, 10:14 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

To be more specific: Can VB6-WebView2 handle async tasks like shown in the C# code?

I'd ask the developer who has programmed the C# interaction (with CEF),
 to write a similar js-function for you...

 WebView2 is using chromium as well, but as said, the DOM-interfaces are not exposed to the VB/COM-side here.
 Therefore you have to use "preloaded js-functions" to interact with the Browser-DOM "on the inside" of WV2 -
 and communicate with the outside via JSON-strings.

 Such a C#->js-v8 porting should not be that hard for the developer, who originally wrote the C#-stuff.

 All I can do here from my end, is to show you the "principle".

Code:

```
Option Explicit

Private WithEvents WV As cWebView2, Elements As cCollection

Private Sub Form_Load()
  Visible = True
  Set WV = New_c.WebView2(hWnd)

  With New_c.StringBuilder 'js-helpers
      .AddNL "function getElements(sExpr){"
      .AddNL "    var resArr = []"
      .AddNL "    for (var e of document.querySelectorAll(sExpr).values()){"
      .AddNL "        resArr.push({tag:e.tagName, id:e.id, val:e.value, name:e.name, href:e.href})"
      .AddNL "    }"
      .AddNL "    return JSON.stringify(resArr)"
      .AddNL "}"

      .AddNL "function setElementBackgroundById(id, color){"
      .AddNL "    document.querySelector('#'+id).style.background = color"
      .AddNL "}"

      WV.AddScriptToExecuteOnDocumentCreated .ToString
  End With

  With New_c.StringBuilder 'two anchors and two buttons as test-HTML-content
    .AddNL "<a id='a1' href='#foo'>LinkFoo</a><br>"
    .AddNL "<input id='b1' type='button' value='button 1'><br>"
    .AddNL "<a id='a2' href='#bar'>LinkBar</a><br>"
    .AddNL "<input id='b2' type='button' value='button 2'><br>"

    WV.NavigateToString .ToString
  End With

  'retrieve a few elements "by selector-expression"
  Set Elements = New_c.JSONDecodeToCollection(WV.jsRun("getElements", "a, input"))

  Dim Elmt As cCollection 'the enumerated Elements themselves are (again) of type cCollection - but hold a JSON-Object instead of a JSON-array
  For Each Elmt In Elements 'iteration over the gathered Element-Objects in the JSON-Array: Elements
     Debug.Print Elmt.SerializeToJSONString 'just to show, which properties are contained "per element"

     Select Case UCase$(Elmt("tag")) 'first Property-access on an Element (here the "tag"-Property)
       Case "A":     WV.jsRun "setElementBackgroundById", Elmt("id"), "#f00"
       Case "INPUT": WV.jsRun "setElementBackgroundById", Elmt("id"), "#0f0"
     End Select
  Next
End Sub
```

... but the adaption of the above shown principle (to your concrete problem), really is your part.

 And as for cairo_sqlite.dll versions:
 - New_c.Cairo.Version ... will hand out the cairo-version
 - New_c.Connection.Version ... will hand out the SQLite-version
 ...and if you put a question-mark before those expression, you can print the results directly in your immediate-window

 HTH

 Olaf

## #313 Schmidt — Mar 27th, 2023, 10:28 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

I have tried to disable the context command, but I think there are no commandline switches available for these settings:

 webView.CoreWebView2.Settings.**AreBrowserAcceleratorKeysEnabled **= false;
 webView.CoreWebView2.Settings.**AreDefaultContextMenusEnabled **= false;

 How could I do this with VB6-WebView2?

The blue marked boolean Properties are exposed on cWebView2
 (as the <F2>reachable VBIDE-ObjectBrowser would show you, or alternatively intellisense, when you type a . behind WV).

 Olaf

## #314 tmighty2 — Mar 27th, 2023, 01:43 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thank you. I access it from outside in an installer to see if I need to deploy a new version to my users.
 There are only a few files that don't provide a version number, and yours were some of them.
 I will work around it.

## #315 tmighty2 — Mar 27th, 2023, 04:17 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The JS operations are really lengthy.
 Would it somehow be possible to make them async and yet return values?

## #316 Schmidt — Mar 28th, 2023, 12:30 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

Would it somehow be possible to make them async ...

Sure... (you're still not familiar with the WV-interface, which you can study in the ObjectExplorer)

 MyToken = WV.jsRunAsync("MyJsFunctionName", MyInputParams, ...)

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

 ...and yet return values?

Private Sub WV_JSAsyncResult(Result As Variant, ByVal Token As Currency, ByVal ErrString As String)
 ' If Token = MyToken Then ...
 End Sub

 HTH

 Olaf

## #317 Bob17 — Mar 28th, 2023, 03:17 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf,

 WebView2 has Stop method, is there a reason why your cWebView2 doesn't have it? Can you add it to the interface?

## #318 Schmidt — Mar 28th, 2023, 04:22 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Bob17** *[img: View Post]*

WebView2 has Stop method, is there a reason why your cWebView2 doesn't have it? Can you add it to the interface?

WV.jsRun "stop" 'should work

 Olaf

## #319 tmighty2 — Apr 1st, 2023, 01:31 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

-----

## #320 tmighty2 — Apr 11th, 2023, 02:46 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I am showing an email in the form of an xml string in the browser.

 Even though the encoding is declared as utf8 and also saved as such, I get strange Umlauts.

 What am I doing wrong?

Code:

```
Public Sub ShowMail(ByVal uXml As String)

    Set WV = New_c.WebView2 'create the instance
    If WV.BindTo(Me.picWV.hwnd, , , , "--autoplay-policy=no-user-gesture-required") = 0 Then
        modLog.WriteLog "#error initilializing browser!"
        MsgBox "couldn't initialize WebView-Binding": Exit Sub
    End If

    WV.AreBrowserAcceleratorKeysEnabled = False
    WV.AreDefaultContextMenusEnabled = False

Dim HtmlFile As String
HtmlFile = New_c.fso.GetLocalAppDataPath & "\temp.html"

New_c.fso.WriteTextContent HtmlFile, uXml, True

WV.Navigate HtmlFile
```

Code:

```
This XML file does not appear to have any style information associated with it. The document tree is shown below.
<mime_message>
<header>
<mime-version>1.0</mime-version>
<date>Fri, 26 Mar 2021 20:38:23 +0100</date>
<message-id> <95B4AC75D4DC194DA3865C264BC6885877F3BB89@xxxxx> </message-id>
<content-type boundary="------------010803010009050608040505"> multipart/alternative </content-type>
<x-priority>3 (Normal)</x-priority>
<from>
<address>
<addr>[email protected]</addr>
<name>xxxxxx</name>
</address>
</from>
<ckx-bounce-address>[email protected]</ckx-bounce-address>
<to>
<address>
<addr>[email protected]</addr>
<name/>
</address>
</to>
<subject>Fwd: AW: </subject>
</header>
<body>
<subpart>
<mime_message>
<header>
<content-type charset="utf-8"> text/plain </content-type>
<content-transfer-encoding>7bit</content-transfer-encoding>
</header>
<body> </body>
</mime_message>
</subpart>
<subpart>
<mime_message>
<header>
<content-type charset="utf-8"> text/html </content-type>
<content-transfer-encoding>quoted-printable</content-transfer-encoding>
</header>
<body>
<![CDATA[ <hr style=3D"color:#62B3FF"><div style=3D"background-color: #DDDDDD; font-s= ize:10pt"><b>Von: </b><a href=3D"mailto:[email protected]">"Anne Willeke= " <[email protected]></a></br><b>Datum: </b>26.03.2021 13:02:16</br><b>A= n: </b><a href=3D"mailto:[email protected]">Johannes xxxxx</a></br><b>B= etreff: </b>Fwd: AW: </div></br><html xmlns:v=3D"urn:schemas-microsoft-com:= vml" xmlns:o=3D"urn:schemas-microsoft-com:office:office" xmlns:w=3D"urn:sch= emas-microsoft-com:office:word" xmlns:m=3D"http://schemas.microsoft.com/off= ice/2004/12/omml" xmlns=3D"http://www.w3.org/TR/REC-html40"> <head><META http-equiv=3D"Content-Type" content=3D"text/html;charset=3Dutf-= 8"> <meta name=3D"Generator" content=3D"Microsoft Word 15 (filtered medium)"> <style><!-- /* Font Definitions */ @font-face {font-family:"Cambria Math"; panose-1:2 4 5 3 5 4 6 3 2 4;} @font-face {font-family:Calibri; panose-1:2 15 5 2 2 2 4 3 2 4;} /* Style Definitions */ p.MsoNormal, li.MsoNormal, div.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;} span.E-MailFormatvorlage18 {mso-style-type:personal-reply; font-family:"Calibri",sans-serif; color:windowtext;} =2EMsoChpDefault {mso-style-type:export-only; font-size:10.0pt;} @page WordSection1 {size:612.0pt 792.0pt; margin:70.85pt 70.85pt 2.0cm 70.85pt;} div.WordSection1 {page:WordSection1;} --></style><!--[if gte mso 9]><xml> <o:shapedefaults v:ext=3D"edit" spidmax=3D"1026" /> </xml><![endif]--><!--[if gte mso 9]><xml> <o:shapelayout v:ext=3D"edit"> <o:idmap v:ext=3D"edit" data=3D"1" /> </o:shapelayout></xml><![endif]--> </head> <body lang=3D"DE" link=3D"#0563C1" vlink=3D"#954F72" style=3D"word-wrap:bre= ak-word"> <div class=3D"WordSection1"> <p class=3D"MsoNormal"><span style=3D"mso-fareast-language:EN-US">Ah, jetzt= habe ich es verstanden, den Fragebogen kannst du auch n=C3=A4chste Woche s= chicken. Klar.<o:p></o:p></span></p> <p class=3D"MsoNormal"><span style=3D"mso-fareast-language:EN-US">Sch=C3=B6= nes Wochenende!<o:p></o:p></span></p> <p class=3D"MsoNormal"><span style=3D"mso-fareast-language:EN-US"><o:p>&nbs= p;</o:p></span></p> <div style=3D"border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0cm = 0cm 0cm"> <p class=3D"MsoNormal"><b>Von:</b> Johannes xxxxx &lt;[email protected]= &gt; <br> <b>Gesendet:</b> Freitag, 26. M=C3=A4rz 2021 09:40<br> <b>An:</b> Anne Willeke &lt;[email protected]&gt;; [email protected]<= br> <b>Betreff:</b> <o:p></o:p></p> </div> <p class=3D"MsoNormal"><o:p>&nbsp;</o:p></p> <p class=3D"MsoNormal"><o:p>&nbsp;</o:p></p> </div> </body> </html> ]]>
</body>
</mime_message>
</subpart>
</body>
</mime_message>
```

*[img: Name:  umlaute.png
Views: 2339
Size:  2.8 KB]*

## #321 Schmidt — Apr 11th, 2023, 04:53 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The "scrambling" has already happened in this case
 (meaning that the passed param uXML already contains "mangled" string-data).

 Please check, where (and how) this uXML-param was retrieved (and then passed) -
 and whether the proper decoding took place whilst filling it elsewhere.

 It might also be, that the xml-header-string (as shown below, for UTF8-mode) -
 points to a code-page, entirely different from UTF8... XML allows that...
 (in which case, you shouldn't write it out as UTF8 via New_c.FSO.WriteTextContent).

 The example code below shows, that when uXML is filled with proper Unicode,
 the WebView2 can show Unicode-XML (in this case with a few german Umlauts), just fine.

Code:

```
Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Visible = True

  Set WV = New_c.WebView2(hWnd)

  Dim uXml$: uXml = "<?xml version='1.0' encoding='UTF-8'?>" & vbCrLf & _
                    "<root><foo>" & ChrW(228) & ChrW(246) & ChrW(252) & _
                    "</foo><bar>" & ChrW(223) & ChrW(8364) & ChrW(248) & "</bar></root>"

  Const FileName = "c:\temp\simple.xml"

  New_c.FSO.WriteTextContent FileName, uXml, True
  WV.Navigate FileName
End Sub
```

HTH

 Olaf

## #322 tmighty2 — Apr 20th, 2023, 02:25 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

----

## #323 tmighty2 — Apr 20th, 2023, 03:30 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

How could I deserialize "Result"?

 This does not compile and throws an Argument Byref Invalid error:

Code:

```
Set Elements = New_c.JSONDecodeToCollection(Result)
```

Code:

```

Private Sub WV_JSAsyncResult(Result As Variant, ByVal Token As Currency, ByVal ErrString As String)

    Debug.Print ErrString

    If Not Token = m_Token Then
        Exit Sub
    End If

    If IsEmpty(Result) Then
       Exit Sub
    End If

'retrieve a few elements "by selector-expression"
    Dim Elements As cCollection
  Set Elements = New_c.JSONDecodeToCollection(Result)
```

This is my JS code:

Code:

```
async function getActiveLinks() {
	let linksArray  = [];

	let linksSource = Array.from(document.querySelectorAll("a, area, button, select, textarea, object, input:not([type=\"hidden\"]), [role=link], [role=button], [role=combobox], [role=option], [role=textbox], [role=checkbox], [role=menuitem]"));

	for (let i = 0; i < linksSource.length; i++) {
		if (await isElelementVisible(linksSource[i], true, true)) {
			let text = await getTextValue(linksSource[i]);
			if (text.length > 0) {
				linksArray.push({
					text : text,
					vsnID: linksSource[i]._vsnEsureId()
				});
			}
		}
	}

		try {
                                          let thestring=JSON.stringify(linksArray);
                                          alert(thestring);
			return thestring;
		} catch (e) {
			alert(e);
		}

}
```

And if I try this...

 Dim s$
 s = Result

 VB6 tells me error 447: "Object does not support the current language setting"

 *[img: Name:  länd2.jpg
Views: 2284
Size:  24.1 KB]*

## #324 Schmidt — Apr 20th, 2023, 05:19 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

How could I deserialize "Result"?

 This does not compile and throws an Argument Byref Invalid error:

Code:

```
Set Elements = New_c.JSONDecodeToCollection(Result)
```

Please try explicit conversion of the Variant Result to a String:
 Set Elements = New_c.JSONDecodeToCollection(**CStr(**Result**)**)

 The VB6-compiler is a bit picky in this case, because **both** -
 Result is a "ByRef" Param and JSONDecodeToCol(...) also expects a ByRef As String argument)

 In other words, VB6 would not have any complaints about passing -
 e.g. a Variant-Return-Result of a VB-**Function** (like e.g. WV.jsRun()) into the JSONDecode-Byref argument directly.

 Here is an example, which works (showing async Result-Handling)

Code:

```
Option Explicit

Private WithEvents WV As cWebView2, Elements As cCollection, AsyncTokens As cCollection

Private Sub Form_Load()
  Visible = True
  Set WV = New_c.WebView2(hWnd)

  With New_c.StringBuilder 'js-helpers
      .AddNL "function getElements(sExpr){"
      .AddNL "    var resArr = []"
      .AddNL "    for (var e of document.querySelectorAll(sExpr).values()){"
      .AddNL "        resArr.push({tag:e.tagName, id:e.id, val:e.value, name:e.name, href:e.href})"
      .AddNL "    }"
      .AddNL "    return JSON.stringify(resArr)"
      .AddNL "}"

      .AddNL "function setElementBackgroundById(id, color){"
      .AddNL "    document.querySelector('#'+id).style.background = color"
      .AddNL "}"

      WV.AddScriptToExecuteOnDocumentCreated .ToString
  End With

  With New_c.StringBuilder 'two anchors and two buttons as test-HTML-content
    .AddNL "<a id='a1' href='#foo'>LinkFoo</a><br>"
    .AddNL "<input id='b1' type='button' value='button 1'><br>"
    .AddNL "<a id='a2' href='#bar'>LinkBar</a><br>"
    .AddNL "<input id='b2' type='button' value='button 2'><br>"

    WV.NavigateToString .ToString
  End With

  Set AsyncTokens = New_c.Collection(False)
  'retrieve a few elements asynchronously "by selector-expression"
  AsyncTokens.Add "getElements", WV.jsRunAsync("getElements", "a, input")
End Sub

Private Sub WV_JSAsyncResult(Result As Variant, ByVal Token As Currency, ByVal ErrString As String)
  If Not AsyncTokens.Exists(Token) Then Exit Sub

  Dim JobType As String: JobType = AsyncTokens(Token): AsyncTokens.Remove Token
  Select Case LCase$(JobType)
    Case "getelements"
      Set Elements = New_c.JSONDecodeToCollection(CStr(Result))

      Dim Elmt As cCollection 'the enumerated Elements themselves are (again) of type cCollection - but hold a JSON-Object instead of a JSON-array
      For Each Elmt In Elements 'iteration over the gathered Element-Objects in the JSON-Array: Elements
         Debug.Print Elmt.SerializeToJSONString 'just to show, which properties are contained "per element"

         Select Case UCase$(Elmt("tag")) 'first Property-access on an Element (here the "tag"-Property)
           Case "A":     WV.jsRun "setElementBackgroundById", Elmt("id"), "#f00"
           Case "INPUT": WV.jsRun "setElementBackgroundById", Elmt("id"), "#0f0"
         End Select
      Next
  End Select
End Sub
```

The other error you'getting is probably related, to not returning a String (but a COM-wrapped js-Object).
 Please ensure, that you also return a String in case of an Error... (from your js-Functions - e.g. via return e.message in the catch-branch).
 It is then relatively easy, to check whether the first char of such a String is either "{" or "[" == a JSON-String ... otherwise an error-string.

 Olaf

## #325 xiaoyao — Apr 20th, 2023, 07:08 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

how to get text from xpath:

Code:

```
Function GetXpathText(Xpath As String) As String
    Dim js As String
    js = "document.evaluate(""" & Xpath & """, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.innerText"
    GetXpathText = Web1.jsProp(js)
End Function
```

Code:

```

 Dim js As String
    js = "document.evaluate(""" & Xpath & """, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.innerText"
 js = "alert(" & js & ")"

Web1.ExecuteScript js
```

alert is ok,but jsProp err,why?

Code:

```
Function GetXpathText(Xpath As String) As String
    Dim js As String
    js = "var XPathText=document.evaluate(""" & Xpath & """, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.innerText"
    Web1.ExecuteScript js
doevents
    GetXpathText = Web1.jsProp("XPathText")
End Function

'Web1.jsProp("XPathText.innerText")
Function GetTextByXpath(Xpath As String, Optional Method As String = ".innerText") As String
    Dim js As String
    js = "var XPathText=document.evaluate(""" & Xpath & """, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
    Web1.ExecuteScript js
doevents
    GetTextByXpath = Web1.jsProp("XPathText" & Method)
End Function
```

Is there a built-in method or more efficient technology that can run this?

## #326 tmighty2 — Apr 21st, 2023, 05:33 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Using my old browser, I defined a callback object (a COM dll) and a sub which would be called by the JS.

 Can you tell me how I would do that with your browser? This is mycode:

Code:

```
async function exportChatAsXMLHelper(params){
	let displayname 	= params[0];
	let includeMedia 	= params[1];
	let debug 			= params[2];

	let xml = await exportChatAsXML(displayname, includeMedia, debug);

	//where is JSCallbackReceiver defined? it is from vb6
	window.JSCallbackReceiver.OnWhatsAppXMLReceived(xml);
}

async function exportChatAsXML(displayname, includeMedia, debug){
	//displayname, includeMedia, debug

    let chat = (await WPP.chat.list()).find(m => m.contact && m.contact.name && m.contact.name.includes(displayname));
    await WPP.chat.openChatBottom(chat.id);
    let msgs = await WPP.chat.getMessages(chat.id, {count : -1});

    const log = (obj) => debug && console.log(obj);

    log('Total messages: ' + msgs.length);

    let xml = '';

    xml+= '<messages>';

    for (var i = 0; i < msgs.length; i++) {
        log('Message number: ' + i);

        let message = msgs[i];
        xml+= '<message>';
        xml+= '<sender>'+ message.from.user +'</sender>';
        xml+= '<receiver>'+ message.to.user +'</receiver>';

        xml+= '<type>'+ (message.type || '') +'</type>';

        if(message.type == 'chat'){
            xml+= '<body>'+ message.body +'</body>';
        }

        if(message.type != 'chat' && includeMedia) {
            xml+= '<media>';
            xml+= '<caption>'+ (message.caption || '') +'</caption>';
            xml+= '<filename>'+ (message.filename || '') +'</filename>';

            log('Downloading media');

            try{
                let mediabody = await mediatoBase64(message.id);
                xml+= '<MediaDownloadStatus>success</MediaDownloadStatus>';
                xml+= '<base64>'+ mediabody +'</base64>';
            }
            catch(e){
                xml+= '<base64></base64>';
                xml+= '<MediaDownloadStatus>fail</MediaDownloadStatus>';
            }
            xml+= '</media>';
        }

        xml+= '</message>';
    }

    xml+= '</messages>';

    return xml;
}

//-----
async function mediatoBase64(msgid) {
    let blob = await WPP.chat.downloadMedia(msgid);
    return new Promise((resolve, _) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.readAsDataURL(blob);
    });
}
```

## #327 Schmidt — Apr 21st, 2023, 09:02 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

Using my old browser, I defined a callback object (a COM dll) and a sub which would be called by the JS.

WV.AddObject is your friend...

 Olaf

## #328 tmighty2 — Apr 21st, 2023, 11:57 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf, I love your work!
 Thank you so much!!

## #329 tmighty2 — Apr 23rd, 2023, 01:02 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Could you tell me how I would call this from VB6?

Code:

```
async function exportChatAsXMLHelper(params){
	let displayname 	= params[0];
	let includeMedia 	= params[1];
	let debug 			= params[2];
//alert(displayname);
	await exportChatAsXML(displayname, includeMedia, debug);
}
async function exportChatAsXML(displayname, includeMedia, debug)
{
    let chat = null;
    if(displayname != ''){
        chat = (await WPP.chat.list()).find(m => m.contact && m.contact.name && m.contact.name.includes(displayname));
    } else {
        chat = (await WPP.chat.list()).find(m => m.active);
    }
    await WPP.chat.openChatBottom(chat.id);
    let msgs = await WPP.chat.getMessages(chat.id, {count : -1});

    const log = (obj) => debug && console.log(obj);

    //alert(msgs.length);

    log('Total messages: ' + msgs.length);

    let count=msgs.length;

    for (var i = 0; i < count; i++) {
        log('Message number: ' + i);

        let message = msgs[i];
        let xml='';
        xml+= '<message>';
        xml+= '<sender>'+ message.from.user +'</sender>';
        xml+= '<receiver>'+ message.to.user +'</receiver>';

        xml+= '<type>'+ (message.type || '') +'</type>';

        if(message.type == 'chat')
        {
            xml+= '<body>'+ message.body +'</body>';
        }

        if(message.type != 'chat' && includeMedia)
        {
            xml+= '<media>';
            xml+= '<caption>'+ (message.caption || '') +'</caption>';
            xml+= '<filename>'+ (message.filename || '') +'</filename>';

            log('Downloading media');

            try
            {
                let mediabody = await mediatoBase64(message.id);
                xml+= '<MediaDownloadStatus>success</MediaDownloadStatus>';
                xml+= '<base64>'+ mediabody +'</base64>';
            }
            catch(e)
            {
                xml+= '<base64></base64>';
                xml+= '<MediaDownloadStatus>fail</MediaDownloadStatus>';
            }
            xml+= '</media>';
        }

        if (i>50)
        {
             //alert('exit!');
           // break;
        }

        xml+= '</message>';

        //where is JSCallbackReceiver defined? it is from vb6
       TheApp.OnWhatsAppXMLReceived(xml,i, count);
       xml='';
    }
  //  xml+= '</messages>';
  //  return xml;
}
```

I have tried this, but it threw an "Object not defined" error:

Code:

```
WV.CallByName Array(sChatName, True, False), "exportChatAsXMLHelper", VbSet

Public Sub OnWhatsAppXMLReceived(ByVal xml As String, ByVal index As Long, ByVal max As Long)

End Sub
```

## #330 Schmidt — Apr 23rd, 2023, 02:23 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

I have tried this, but it threw an "Object not defined" error:

Code:

```
WV.CallByName Array(sChatName, True, False), "exportChatAsXMLHelper", VbSet

Public Sub OnWhatsAppXMLReceived(ByVal xml As String, ByVal index As Long, ByVal max As Long)

End Sub
```

Not sure, why you're using WV.CallByName, when it was already established,
 that js-functions shall be called via either:
 - WV.jsRun( "someJsFunction", ..., ... ) or
 - WV.jsRunAsync "someOtherFunc", ..., ...

 Also cannot deduce the purpose of your intermediate function (exportChatAsXMLHelper) -
 you can call the main-routine perfectly fine (with its 3 arguments) via e.g. WV.jsAsync...

 Here is an example again:

Code:

```
Option Explicit

Private WithEvents WV As cWebView2, Elements As cCollection

Private Sub Form_Load()
  Visible = True
  Set WV = New_c.WebView2(hWnd)
      WV.AddObject "TheApp", Me

  With New_c.StringBuilder
      .AddNL "async function exportChatAsXML(displayname, includeMedia, debug){"
      .AddNL "    TheApp.ReportIncomingArguments(displayname, includeMedia, debug)"
      .AddNL "}"

      WV.AddScriptToExecuteOnDocumentCreated .ToString
  End With
  WV.NavigateToString "Dummy-Document-Content"

  WV.jsRunAsync "exportChatAsXML", "TheDisplayName", True, False
End Sub

Public Sub ReportIncomingArguments(sDisplayname, bIncludeMedia, bDebug)
  Debug.Print sDisplayname, bIncludeMedia, bDebug
End Sub
```

Olaf

## #331 tmighty2 — Apr 23rd, 2023, 03:44 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The error that I'm getting is not really related to the browser, I guess, but I wanted to ask anyway.

 <https://drive.google.com/file/d/1WwW9MFwS6_qYu3oTIjgaYdR1G1nl-CZt/view?usp=sharing> https://drive.google.com/file/d/1WwW...ew?usp=sharing

 I am injecting the script using AddScriptToExecuteOnDocumentCreated.

 I have successfully done this:

 Public Sub ReportIncomingArguments(sDisplayname, bIncludeMedia, bDebug)
 Debug.Print sDisplayname, bIncludeMedia, bDebug
 End Sub

 However, when I try to use my WhatsApp function, I get an error saying

 VM11:2 Module Conn was not found with e=>e.Conn&&e.ConnImpl
 get @ VM11:2
 VM11:2 Uncaught (in promise) TypeError: Cannot read properties of undefined (reading 'on')
 at E.<anonymous> (<anonymous>:2:164253)
 at E.emitAsync (<anonymous>:2:53133)
 at <anonymous>:2:239794
 at a (runtime.276e78425ec3861add7e.js:1:15161)
 at Array.forEach (<anonymous>)
 at runtime.276e78425ec3861add7e.js:1:15333
 at runtime.276e78425ec3861add7e.js:1:15393
 at runtime.276e78425ec3861add7e.js:1:15397

 What could I check to find out where it goes wrong?

 The JS file was done by somebody else. I don't know where the webpack code came from.
 But I know that I successfully used the JS in my old browser.

 If anybody can suggest what to try to pinpoint the problem, I would be glad.
 Thank you!

## #332 Schmidt — Apr 23rd, 2023, 08:29 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

The error that I'm getting is not really related to the browser, I guess, but I wanted to ask anyway.

 <https://drive.google.com/file/d/1WwW9MFwS6_qYu3oTIjgaYdR1G1nl-CZt/view?usp=sharing> https://drive.google.com/file/d/1WwW...ew?usp=sharing

 ...
 However, when I try to use my WhatsApp function, I get an error saying

 VM11:2 Module Conn was not found with e=>e.Conn&&e.ConnImpl

Looks like:
 - another js-dependency (or js-module) is still missing (not pre-loaded)...
 - or your larger js-block runs "too early" (try loading it in WV_DocumentComplete with WV.ExecuteScript instead)
 - or the script-code needs to be run from within a script tag with attribute type='module': <https://stackoverflow.com/questions/65657954/must-all-script-tags-be-type-module-when-using-module-syntax> https://stackoverflow.com/questions/...-module-syntax

 Best advice I can give at this point, is to get yourself help from a js-Expert...

 Olaf

## #333 tmighty2 — Apr 24th, 2023, 03:00 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I did that, and it turned out that the webpack code was either outdated or malformed in some way.
 Olaf, your components help me to pay my living since 20 years :-)

## #334 tmighty2 — Apr 26th, 2023, 09:16 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

If I have added code using "AddScriptToExecuteOnDocumentCreated", how would I clear that again?

 I have not checked, but I am afraid that some parts of the JS may be double as I add several JS files.
 And some websites do not need all of these scripts.

 That is why I wonder if I should do housekeeping and keep everything clean.

## #335 Schmidt — Apr 26th, 2023, 02:37 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

If I have added code using "AddScriptToExecuteOnDocumentCreated", how would I clear that again?

 I have not checked, but I am afraid that some parts of the JS may be double as I add several JS files.
 And some websites do not need all of these scripts.

 That is why I wonder if I should do housekeeping and keep everything clean.

As already mentioned in my last post - an alternative way to add script-modules,
 is to do it in the DocumentComplete-Event... via WV.ExecuteScript your-module-source-code.

 Within that event you can react conditionally what to load, depending on the current Document-URL.

 AddScriptToExecuteOnDocumentCreated is useful for:
 - truly generic, never-changing helper-functions - which work and are useful on any site
 - or to cover very specific "local" or "single-site" scenarios
 ...(using the current WV-instance which was placed on a specific Form).

 Also keep in mind that, when you e.g. write your own UserControl -
 you can have different WV-instances (each of them bound to the internal UserControl.hWnd)...
 Such a control could have its own Public URL-Property Let and Get ...
 and multiple instances of such a UC could be arranged for e.g. "tabbed Browsing".
 You could decide, what dependencies you will load - by looking at the incoming NewURL-Param in the URL-Prop-Let-routine.

 Olaf

## #336 tmighty2 — May 1st, 2023, 04:02 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf, are you aware of that RemoteMessenger.postSyncRequestMessage error popping up in the DevToolsWindow each time?

## #337 Schmidt — May 2nd, 2023, 05:27 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

Olaf, are you aware of that RemoteMessenger.postSyncRequestMessage error popping up in the DevToolsWindow each time?

Yes, it will be fixed in the next release.

 Olaf

## #338 xiaoyao — May 2nd, 2023, 11:58 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

......

## #339 Schmidt — May 3rd, 2023, 04:36 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

can't open url:
 <https://www.365365091.com/#/IP/B13> https://www.365365091.com/#/IP/B13

Please explain yourself properly (instead of posting dubious links to some "betting-websites").

 Olaf

## #340 tmighty2 — May 6th, 2023, 08:03 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Would it have any bad impact if I have 2 instances of the browser?
 I am not sure if there is a way to make something like tabs.
 I would like to quickly switch from etc. google.com to youtube.com.
 I manipulate the DOM and execute scripts on these websites. I am currently stuck because my google.com scripts and youtube.com scripts interfere. I guess I can make that work, but I would anyways like to ask if it woul be good practice to keep these 2 apart by using 2 browser instances.

## #341 xiaoyao — May 6th, 2023, 10:57 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

It is best to create two processes, otherwise if two web pages run in one thread, the speed will become slower

## #342 Schmidt — May 6th, 2023, 12:34 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

It is best to create two processes, otherwise if two web pages run in one thread, the speed will become slower

The chromium-engine (which MS-Edge-WebView2 is using under the covers) -
 is already working "multi-process-based" underneath - a short view into the Taskmanager would show you that...

 So, the **hWnd **a Client-Process (the VB6-App) will bind a WV-instance to,
 is only the "tip of the iceberg" (a ViewPort-rectangle) - these multiple chromium-processes finally "Blit-to" -
 (and in the opposite direction, will only receive KeyBoard- and Mouse-Messages from)...

 So, no - there is no performance-penalty (and also no memory-problem) with multiple WV-instances in a given VB-App.

 Besides, could you please post only about topics you're really knowledgeable about?
 (and if you do, then in "better understandable english").

 Currently you come across as a "spamming lunatic", resembling a "badly implemented chat-bot".

 Olaf

## #343 saturnian — May 6th, 2023, 01:33 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

 Currently you come across as a "spamming lunatic", resembling a "badly implemented chat-bot".

 Olaf

that is perfectly correct!
 you had to dare to say it! Thanks Olaf for making it !

 *[img: big yellow]*

## #344 Bob17 — May 8th, 2023, 05:00 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,
 you have the method "OpenDevToolsWindow", could you please also add "CloseDevToolsWindow" to the API?

## #345 Schmidt — May 8th, 2023, 09:36 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Bob17** *[img: View Post]*

Hi Olaf,
 you have the method "OpenDevToolsWindow", could you please also add "CloseDevToolsWindow" to the API?

AFAIK, there's currently no built-in method in the official typelib for that...

 So you have to go via "normal API" (e.g. using FindWindow - the Window seems always to start with the Caption-prefix "DevTools").

 HTH

 Olaf

## #346 paliadoyo — May 8th, 2023, 03:10 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,
 How to use the "SetFuncObj" method?

 Thanks!

## #347 Schmidt — May 8th, 2023, 04:54 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **paliadoyo** *[img: View Post]*

How to use the "SetFuncObj" method?

It's a method, meant for "RC6-internal-use" (to establish easier communication with the js_side of the WV2 in the initialization-phase) -
 and I had to make it a Public-Method instead of Friend, for OLE-marshaling to work...)

 Once a WV is initialized, the "official UserApi" to place COM-instances in the js-context (e.g for callback-purposes), is:
 WV.AddObject ..., ...

 Olaf

## #348 tmighty2 — May 8th, 2023, 05:45 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Is there any way to block this error from occuring until the new version has been released? It is pretty distracting in the dev console window.

## #349 paliadoyo — May 9th, 2023, 02:08 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

It's a method, meant for "RC6-internal-use" (to establish easier communication with the js_side of the WV2 in the initialization-phase) -
 and I had to make it a Public-Method instead of Friend, for OLE-marshaling to work...)

 Once a WV is initialized, the "official UserApi" to place COM-instances in the js-context (e.g for callback-purposes), is:
 WV.AddObject ..., ...

 Olaf

Perfect. Thanks!

## #350 tmighty2 — May 9th, 2023, 10:51 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Under some circumstances, "WV.DocumentURL" might take around 3 seconds.

 If I get it in WV_DocumentComplete like this...

Code:

```
Private Sub WV_DocumentComplete()

    g_sCurrentURL = WV.DocumentURL
    g_sCurrentTitle = WV.DocumentTitle
```

... it is always fast.
 However, when I call it at other times, it might take up to 3 seconds. That is why I store it in WV_DocumentComplete.

 Thanks!

## #351 tmighty2 — May 10th, 2023, 01:41 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I can execute JS on "https://google.de" or "https://google.com.tr" just fine.
 On "https://google.com", it does not work.
 Is anybody familiar with this?
 I will post a sample project shortly.

## #352 tmighty2 — May 10th, 2023, 01:51 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I can execute JS on "https://google.de" or "https://google.com.tr" just fine.
 On "https://google.com", it does not work.
 Is anybody familiar with this?

 Here is a sample project: <https://drive.google.com/file/d/1aYjqxsIKdMpMfWLbKfOMzGXUPp1IKsu0/view?usp=sharing> https://drive.google.com/file/d/1aYj...ew?usp=sharing
 And here is a video: <https://www.youtube.com/watch?v=Snz8JN6zGw4> https://www.youtube.com/watch?v=Snz8JN6zGw4

## #353 tmighty2 — May 10th, 2023, 11:18 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Code:

```
window.addEventListener ? (document.addEventListener("DOMContentLoaded", c, !1), window.addEventListener("load", c, !1)) : window.attachEvent && window.attachEvent("onload", c);
```

in <https://google.com> https://google.com

 This broke webview2 host object and other webview2 functions

Code:

```
	(function () {
		var b       = [function () {
			google.tick && google.tick("load", "dcl");
		}];
		google.dclc = function (a) {
			b.length ? b.push(a) : a();
		};

		function c() {
			for (var a = b.shift(); a;) a(), a = b.shift();
		}

		//window.addEventListener ? (document.addEventListener("DOMContentLoaded", c, !1), window.addEventListener("load", c, !1)) : window.attachEvent && window.attachEvent("onload", c);
	}).call(this);
```

## #354 tmighty2 — May 11th, 2023, 03:33 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

This morning I woke up wondering if Olaf Schmidt has a Wikipedia entry.

## #355 carl039 — May 12th, 2023, 09:38 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

**Found cause to following issue, no reply needed.**

 Hi

 I am running into issue with Windows Server 2016 which can't even run the WebView2Demo which results in the following error on launch:

 Run-time error '430':

 Class does not support Automation or does not support expected interface.

 The same version of RC6 controls/dll's on my Windows 10/11 workstations work fine but not on 2016.

 Have reinstalled MS Webview2 components and also installed MS Edge Browser to but that didn't help.

 Any ideas on what I am missing?

 Thanks

 Carl

## #356 Schmidt — May 12th, 2023, 01:57 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **carl039** *[img: View Post]*

**Found cause to following issue, no reply needed.**

 ...I am running into issue with Windows Server 2016 which can't even run the WebView2Demo which results in the following error on launch:

Just out of interest, was it a missing "registration" of the RC6.dll on that server-machine?
 (or something else)?

 Olaf

## #357 xiaoyao — May 13th, 2023, 06:35 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **carl039** *[img: View Post]*

**Found cause to following issue, no reply needed.**

 Hi

 I am running into issue with Windows Server 2016 which can't even run the WebView2Demo which results in the following error on launch:

 Run-time error '430':

 Class does not support Automation or does not support expected interface.

 The same version of RC6 controls/dll's on my Windows 10/11 workstations work fine but not on 2016.

 Have reinstalled MS Webview2 components and also installed MS Edge Browser to but that didn't help.

 Any ideas on what I am missing?

 Thanks

 Carl

Official: Microsoft Edge WebView2 profile - Microsoft Edge Development | Microsoft Learn
 <https://learn.microsoft.com/zh-cn/microsoft-edge/webview2/> https://learn.microsoft.com/zh-cn/mi...edge/webview2/
 WebView2 runtime version 109 is the final version that supports the following versions of Windows. The WebView2 runtime and SDK version 110.0.1519.0 and later do not support these operating systems.

 Windows 8/8.1
 Windows 7
 Windows Server 2012 R2
 Windows Server 2012
 Windows Server 2008 R2

## #358 xiaoyao — May 15th, 2023, 02:11 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I heard that there is a way to read and write MYSQL database without installing the driver, just extract a DLL like SQLITE3.DLL?
 Maybe a special ADOx86.DLL,ado x64.dll? Can I read and write to the database without installation?
 This way, different operating systems will have the same experience and won't cause all kinds of compatibility issues and errors.
 EDGE, for example, has a version 109 that supports older Windows 7, Windows 8, and so on. But he can't have a single DLL for all operating systems.
 MINIBLINK.DLL does this (a DLL is only 17m-35M, it's like installing IE11, an old version of Google Chrome, some of which is not supported, but works perfectly with electron,node js), but the latest version is only blink 75. Now Google is in version 113

## #359 xiaoyao — May 15th, 2023, 06:22 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

i want to show webview by vb.net with rc6.dll

 In the SHARPDEVELOP developer tool, the reference to RC6.DLL does not show the class
 Unable to load type library referencing "RC6". The library is not registered. (Exception from HRESULT:0x8002801D (TYPE_E_LIBNOTREGISTERED)) (MSB3287)
 Is your RC6.DLL developed in VC++2010? It took me a lot of time to install the WINDOWS 8.0SDK before I could compile it

 HOW TO USE IN VB.NET?LIKE THIS

 Next you need to register the DLL. Select the Start menu's Run command and execute the statement:

 regsvr32 VB6Project.dll
 Next start a Visual Basic .NET project. Select the Project menu's Add Reference command. Click the COM tab and find the DLL or click the Browse button to select it. Now the .NET application can use the DLL's public classes as shown in the following code.

Code:

```
Private Sub btnCallSubroutine_Click(ByVal sender As _
    System.Object, ByVal e As System.EventArgs) Handles _
    btnCallSubroutine.Click
    Dim vb6_class As New VB6Project.MyVB6Class
    vb6_class.VB6SayHi()
End Sub

Private Sub btnCallFunction_Click(ByVal sender As _
    System.Object, ByVal e As System.EventArgs) Handles _
    btnCallFunction.Click
    Dim vb6_class As New VB6Project.MyVB6Class
    MessageBox.Show("Returned: " & vb6_class.VB6ReturnHi())
End Sub
```

## #360 Shaggy Hiker — May 16th, 2023, 09:15 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Those last two posts do not appear to be related in any way to the topic at hand. The topic does seem to have wandered a bit, so the second one is kind of on topic, but not the post about SQLLite.

## #361 xiaoyao — May 19th, 2023, 03:44 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Found two problems, help to have a look.

 webview2 loading objects is slow, can you do an asynchronous version, as long as you set the URL, the background will automatically create, create an event after completion. And then go to the website

HTML Code:

```
Set Webview1 = New_c.WebView2
Webview1.BindTo(UserControl.hWnd)
```

Code:

```
Set Webview1 = New_c.WebView2
webview1.NewWebViewAnsy(picture1.hwnd,URL AS STRING)

SUB NewWebViewAnsy(ParentHwnd as long,Url1) IN CreateThread or ansy
Webview1.BindTo(UserControl.hWnd)
raiseevent  WebviewLoadOK()
Webview1.Navigate URL1
end sub
```

Url without http added will Navigate and crash Webview1.Navigate "baidu.com

## #362 Schmidt — May 19th, 2023, 07:52 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

webview2 loading objects is slow, can you do an asynchronous version

Async mode is already built-in, via the SecondsToWaitForCompletion-parameter.
 (in the Constructor/Bind and also in the Navigate-Methods).

 There is also an Event, which signalizes a successful binding (the InitComplete-Event).
 Here is an example, how to work asynchronously (with constructor-binding and navigation):

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Set WV = New_c.WebView2(hWnd, 0) '<- setting the "TimeToWait" to 0, will do this bind-call asynchronously

  Debug.Print "now exiting Form_Load (relying on WV_InitComplete)"
End Sub

Private Sub Form_Resize()
  If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub

Private Sub WV_InitComplete()
  Debug.Print "InitComplete"
  WV.Navigate "https://google.com", 0 '<- same 0-TimeToWait for an async Nav-method-call
End Sub
```

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

Url without http added will Navigate and crash Webview1.Navigate "baidu.com

It does not "crash" - instead what it does, is raising an Error 5 (wrong argument).

 And a wrong argument means (in this case), that the URL-Param requires a "protocol-prefix"
 (like "http://" or "https://")

 Olaf

## #363 xiaoyao — May 19th, 2023, 10:01 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

WV_InitComplete()
 This is a good way.
 If at the beginning of the binding object, directly set the URL to complete the asynchronous do not have to pay attention to, so better.

 The radio parameters are done automatically as if it were a multi-threaded delegate, and then notify me.I mainly use it to load 10 webview controls at the same time in one second, and each control opens the URL according to the array, so I just need to send out the task and he will finish it in the background.

## #364 xiaoyao — May 19th, 2023, 10:09 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

It would be better if 64-bit rc6.dll can be developed. Only some functions can be realized. Can you support the function of creating websocket server, so that I can directly open Google browser and use the plug-in ws://127.0.0.1 Connect to the RC6 point dll. I don't know if webview2 object supports edge and Google plug-in. If not, I will load a js URL or local script. Pytho operates Google and edge driver, and they provide more functions and are more convenient. So if you use this, there is still a lot of code to be implemented in js.

 rc6.DLL，Is it possible to remove sqlite3.dll if you are not using a database?

## #365 xiaoyao — May 19th, 2023, 12:27 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

To test, I wrote a multi-tab browser, which begins to work well.
 François

Have you ever developed a custom control to wrap webview2?
 So you can use it like webbrowser. I have encountered many puzzling problems, such as the size has not changed and I can't display the web page.Is there any way to operate in an Excel vba,vbs file? It will automatically open an edge browser window or a simple form display page. You don't need to bind to a handle. And if you click on a web page, it will pop up a new tab or page. Can you also read its contents with code?

## #366 softv — May 20th, 2023, 11:03 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

If I have added code using "AddScriptToExecuteOnDocumentCreated", how would I clear that again?

 ... .. .

Dear Olaf,

 I also need a way to remove the script added via "AddScriptToExecuteOnDocumentCreated". Thanks a lot for your answers in connection with this but I did not find them helping me in my case. So, I tried my very best (for the past 3 days or so, whenever I could get time) to find a way out (some sort of hack at least) to tackle this issue but I am unable to, so far. So, I am stuck.

 Today, though, when browsing the net in connection with the above issue, I came across this page - <https://learn.microsoft.com/en-gb/microsoft-edge/webview2/reference/win32/icorewebview2?view=webview2-1.0.1774.30#removescripttoexecuteondocumentcreated> https://learn.microsoft.com/en-gb/mi...ocumentcreated ( **RemoveScriptToExecuteOnDocumentCreated **). Reading what this function does, it seems like this is what I require. But then, you are the expert. May be the aforesaid function does NOT exactly do what I require. If so, kindly bear with me and kindly let me know the same. On the other hand, if at all the aforesaid function indeed does what I require, then I request you to kindly include it in RC6 WV, **if at all possible**. It will simplify my tasks like anything. I would be much grateful to you for the same.

 I take this opportunity to once again **thank you in tons** for all your contributions and support.

 Kind Regards.

## #367 xiaoyao — May 22nd, 2023, 11:16 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

In addition, my current project is almost finished. In a few months, I will start to develop a scripting language and IDE similar to VB6 (of course, the development time is also calculated in years). If you could provide some good seeds (prototypes or suggestions), maybe I can develop a decent scripting language and IDE, and this IDE may be applied to your VB6-compatible compiler.

 (From dm)

I don't know how it goes now. Where can I download the preview version?

## #368 xiaoyao — May 22nd, 2023, 11:43 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

- then instantiate New_c regfree within your App (via a declared DirectCOM.dll GetInstance or GetInstanceEx-call)
 - and then initiate the Binding by a Class-instance, derived via New_c.WebView

 1，rc6.dll
 2，sqlite3.dll
 3，webview2load.dll
 4，DirectCOM.dll
 5，abc.exe >> bind to vb6 picture1.hwnd

 This requires so many files that I don't mind the file size.
 DirectCOM.dll,In fact, the output API can be implemented directly in rc6.dll. The ActiveX DLL object can also output the stdcall API,
 If SQLite 3.dll can specify the path, it would be better to modify the file name. Register rc6.dll without sqlite3.dll, I just load it when I need it, that's fine.

 An interface can be added,Display a web page window directly without writing any code.
 like this in vb6

Code:

```
Set ie = CreateObject("InternetExplorer.Application")
ie.Navigate "http://www.baidu.com"
ie.visible=true
```

quick open url on excel vba,or vbscript:

Code:

```
Public Declare Function DirectCom_Create Lib "rc7.dll" Alias "GETINSTANCE" (FName As String, className As String) As Object
Public Declare Function NewWindow Lib "rc7.dll"   (byval url As String) As Object

dim web2 as object
set web2=DirectCom_Create("rc7.dll","cwebview2")
web2.NewWindow("http://www.baidu.com")
web2.visible=true
```

or

Code:

```
dim web2 as object
set web2=NewWindow("http://www.baidu.com")
web2.visible=true
```

like new version of rc7.dll
 rc7.vpb add text to end:
 [VBCompiler]
 LinkSwitches=-EXPORT:NewWindow -Export:GetInstance -Export:AutoSetup -Export:CheckSetup -Export:Add -EXPORT:ShowForm1

Code:

```
Dim Web1 As CwebView2

Public Function NewWindow(Url As String) As CwebView2
    Set Web1 = New CwebView2
    Web1.Navigate Url
    Set NewWindow = Web1
End Function

Public Function GetInstance(FName As String, className As String) As Object
    ' code
End Function

Public Function AutoSetup() As Boolean
    'download Url = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
    'download Url to Edge_Webview2RunTime.exe
     'ShellWait(App.Path & "\Edge_Webview2RunTime.exe  /silent /install"
End Function

Public Function CheckSetup() As Boolean
'
End Function
```

## #369 tmighty2 — May 23rd, 2023, 05:06 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

We are trying to create an adblocker.
 Here is what my collegue would like to ask:

 Is it possible to use AddWebResourceRequestedFilter method?
 I can't write to form, it says I am not authorized.
 Can you give an example about using AddWebResourceRequestedFilter method, if possible?

## #370 xiaoyao — May 23rd, 2023, 10:43 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Want to turn rc6.dll Cwebview2 objects into a portable browser, without any ads and other menus
 You can call the

Code:

```
rundll32.exe rc6.dll,WebViewOpenurl,https://www.163.com
```

with the vbs
 Methods like this. This displays the form interface in the dll and loads a web page

 Features that can be achieved, automatically maximize the form, or hide the top navigation bar of the form. Pure web pages can also be mobile.
 Set the form size, remember the form size, the next time open or the same location.
 Main purpose, quickly open some web pages, or form the top way to load web page applications. Like a mobile web-page calculator.
 Or are shown as the front end of the PYTHON. PYTHON The back-end uses the websocket service to receive the data, enter the web page order, and deposit PY in the database.

 You can directly use VBS and EXCEL VBA to load web pages to obtain objects, control webview2 objects, send JS commands, and collect web page content.

Code:

```
SET WEB=createobject("rc6.cwebview2")
web.WindowOpen("https://www.baidu.com")
web.visiable=true
```

## #371 Schmidt — May 24th, 2023, 07:03 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

Want to turn rc6.dll Cwebview2 objects into a portable browser...

Why use "rundll.exe", when you can compile a small executable with VB6 - and make it load the New_c in a regfree manner -
 via the already outlined methods, as e.g. here: <https://www.vbforums.com/showthread.php?897817-Registry-Free-Object-Instantiation-using-DirectCOM-amp-RC6> https://www.vbforums.com/showthread....ectCOM-amp-RC6

 Olaf

## #372 tmighty2 — May 28th, 2023, 06:10 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The problem is that we are missing the WebResourceRequested event.
 Event WebResourceResponseReceived(ReqURI As String, ReqMethod As String, RespStatus As Long, RespReasonPhrase As String, RespHeaders As cCollection) is already too late, I think.

## #373 Schmidt — May 29th, 2023, 01:57 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

The problem is that we are missing the WebResourceRequested event.
 Event WebResourceResponseReceived(...) ...is already too late, I think.

WebResourceResponseReceived is triggered from e.g.:
 WV.NavigateWithWebResourceRequest

 You're right in, that the proper Event which corresponds with: WV.AddWebResourceRequestedFilter -
 is the WebResourceRequested which is now included (in just uploaded RC6-version 6.0.13)...

 I've also added the new NavigationStarting-Event (which is another one, which allows cancelling requests - on "whole documents").

 Here's how to use the new WebResourceRequested-Event (which will only be triggered, once a Filter was added):

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Visible = True

  Set WV = New_c.WebView2(hWnd)

  WV.AddWebResourceRequestedFilter "*googlelogo*" 'define an URI-filter with wildcards

  WV.Navigate "https://google.com"
End Sub

Private Sub Form_Resize()
  If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub

'requires RC6 >= 6.0.13
Private Sub WV_WebResourceRequested(ByVal URI As String, ByVal FilterType As RC6.eWebView2ResourceFilter, Permit As RC6.eWebView2PermissionState)
  Debug.Print FilterType, FilterType = Filter_IMAGE, URI
  Permit = PERMISSION_STATE_DENY
End Sub
```

So, the above will report (in the Event), all "sub-requests" which in their URI contain the "wildcarded" string-snippet
 (in case of the example, the filter-string was "*googlelogo*").

 Within the Event you can then take a closer look at the URI which contained that filter-snippet -
 and then decide (depending on resourcetype and "whole URI-string", whether you want to "403 cancel" it, or not).

 HTH

 Olaf

## #374 softv — Jun 3rd, 2023, 09:00 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Dear Olaf,

 In post no. <https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)&p=5604638&viewfull=1#post5604638> 342, you have written as follows:
 // The chromium-engine (which MS-Edge-WebView2 is using under the covers) -
 is already working "**multi-process-based**" underneath - a short view into the **Taskmanager** would show you that... //

 Thanks for the above info and I could indeed see the MS-Edge-WebView2 process(es) in TaskManager. I have a few questions related to it. As far I searched through this thread, I could not find these questions asked. Hence, asking.
 ----------
 Suppose I have an app (A.exe) showing a web page in its 1st screen (say 'X') using WV.

 When I open A.exe, some "Microsoft Edge WebView2" processes are shown running in the Processes list. If I close A at once, all the aforesaid "Microsoft Edge WebView2" processes disappear from the Processes list.

 Instead of the above scenario, after opening 'A', if I click a link in 'X' (which link has **target="_blank"** and it is for a web page 'Y' ) so that 'Y' web page opens in a new window (say **'N'**). Now, when I close 'A', 'N' window showing 'Y' web page still remains on the screen. Upon checking the Processes list, there remains in it a primary "Microsoft Edge WebView2" process with some child processes inside it. Is this the normal behaviour? If so, what exactly is happening? I presume 'N' showing 'Y' is running now as a separate process in a separate "Microsoft Edge WebView2" runtime process (say **'M'**). Am I right?

 Well, in any case, based on the abovementioned still-running 'N' window, I have the following sub-questions:

 1. What is the best way to automatically **close such 'N' windows** (one or many) when I close my main application 'A'? Is there a possibility to have a property in WV by which I say that whenever I close 'A', all 'N' windows opened by 'A' shall be closed automatically? Or, is there already present a simple method in WV by which during Unload of 'A', I can close all 'N's easily?

 2. What if I **don't close such 'N' windows** and open my application 'A' again? I presume there will be no relationship between the already-opened 'N' windows and 'A'. Am I right in this presumption? If so, if I allow users to open many such 'N' windows in many different sessions of 'A', then each one of those 'N' windows will be running under the abovementioned **'M'** process (of "Microsoft Edge WebView2") and user can opt to close each of those 'N' windows himself, whenever he wishes to, later. Right? In other words, there will not be any problems at all in 'A' being opened/closed several times when many already opened 'N' windows are existing. Right? **Kindly let me know**.
 ----------

 I take this opportunity to **thank you once again**, heartily, from the bottom of my heart, for all that you do to help us out.

 Kind regards.

## #375 xiaoyao — Jun 5th, 2023, 04:33 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I INSTALL this,MS-DownloadLink
 <https://go.microsoft.com/fwlink/p/?LinkId=2124703> https://go.microsoft.com/fwlink/p/?LinkId=2124703

 Error in FindFirstFile: ErrNum: 3, The system cannot find the specified path

 OK = Web1.BindTo(Web1pic.hWnd) <> 0

 ==========
 I specify the old version and it works：
 OK = Web1.BindTo(Web1pic.hWnd, , App.Path & "\SystemFile\web110.0.1587.63.x86")

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

There seems to be a problem with the installation of the WebView-Runtime...

 The BindTo-call is looking for the WebView-Runtime-BaseFolder via the following Path-Expression:
 (for your Immediate-Window):
 ?New_c.FSO.GetSpecialFolder(CSIDL_PROGRAM_FILESX86) & "\Microsoft\EdgeWebView\Application"

 Olaf

i search in hrer:C:\Users\***\AppData\Local\Microsoft\EdgeWebView
 not C:\Program Files (x86)\

Code:

```
HKEY_CURRENT_USER\Software\Microsoft\EdgeUpdate
path=C:\Users\xiaoyao\AppData\Local\Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe
```

how to get newwindow by webview2 control?
 and how to do that new webvie2 like runjs,openurl(**)

 if have 3 new window ,how to do?

## #376 xiaoyao — Jun 5th, 2023, 04:49 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

last time ,i found someone say check reg path from :HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}\
 cant find location in this regpath

 It is possible that the Webview sdk is not installed with administrator permission. Therefore, there is no data in HKEY_LOCAL_MACHINE. Only HKEY_CURRENT_USER is saved

 HKEY_CURRENT_USER\Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}

 edgesdk=""
 CreateCoreWebView2EnvironmentWithOptions(StrPtr(edgesdk), StrPtr(userdata)

 A direct API call with an empty path will automatically detect where the system is installed.

## #377 xiaoyao — Jun 5th, 2023, 05:29 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

As for detection, whether a given target-system already contains this runtime,
 you can simply check the output of the GetMostRecentInstallPath-method:
 This function failed to get the WEBVIEW SDK installation path

 Set Web1 = New_c.WebView2
 Web1.GetMostRecentInstallPath

 RC6.DLL 2023-5-29 version still has this problem

Code:

```
    Dim OK As Boolean
    Set Web1 = New_c.WebView2
    Dim WebViewSDK As String
    WebViewSDK = GetEdgeWebViewPath
    If WebViewSDK <> "" Then
        OK = Web1.BindTo(Me.hWnd, , WebViewSDK) <> 0
    End If

Function GetEdgeWebViewPath() As String
   On Error Resume Next
   Dim WshShell As Object, WebView2_Version As String
   Dim RegPath As String, location As String
   Set WshShell = CreateObject("Wscript.Shell")
   RegPath = "HKEY_CURRENT_USER\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}\"
   location = WshShell.RegRead(RegPath & "location")
   If location = "" Then
        RegPath = Replace(RegPath, "\WOW6432Node", "")
        location = WshShell.RegRead(RegPath & "location")
   End If
   If location <> "" Then GetEdgeWebViewPath = location & "\" & WshShell.RegRead(RegPath & "PV")
End Function
```

## #378 xiaoyao — Jun 5th, 2023, 09:09 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

In my understanding, Selenium was written as a generic Test-Tool which can "work against multiple Browser-engines".

 The WebView2 is just an encapsulation of a single Browser-engine (chromium) -
 minus the additional "generic Automation- and Test-interfaces".

 So, if you currently use Selenium for your own testing-purposes with success,
 why change it to something "less capable" and "less generic".

 Just my $0.02...

 Olaf

I used to do this with IE, and there was a program that would automatically start IE and add some command line arguments.
 So I used VB6 to write a program, replace IE path, and then use the WEBBROWSER to load the web page, the command line needs to open the URL.
 chrome.exe works the same way. chromedriver.exe will send a WebSocket-related URL, and then through developer tools and other principles, you can use the driver proxy to open the URL, click the button.

 I went over everyone's responses and questions several times today. And tested it.
 If everyone can solve the problem, then do a test code is convenient. Then indicate which forum the question is placed in.

 I heard that Google Chromium kernel source code related projects have a total of about 20G, the final compilation of DLL and related files may only be several hundred M.
 Just because of the amount of documentation, the development tools, are important assets, learning paths.

 I've always been studying very seriously, but sometimes the focus is different.
 FOR EXAMPLE, I WANT RC6 .DLL TO RUN INDEPENDENTLY, SO THAT A WEB PAGE CAN BE DISPLAYED BY CALLING A STANDARD DLL API AND A WEBVIEW OBJECT CAN BE OBTAINED.

 It's like VB6 calling WebView2Loader.dll, which is a VC++ COM object, but also outputs the 4 API exports of the standard DLL.
 With this, there is a complete source code and webview2.tlb, which can be added in VB6.
 So if RC6 .DLL can expose webview objects, then many users can extend some of the functionality that RC6 lacks.
 If you only need an API to directly display a form loading web page by yourself, you will not cause multithreaded calls to some callbacks in the VB.NET and crash.

 Of course, some features are added that are really convenient, but they also take a lot of time and technical difficulty to do.
 After 25 years, Twinbasic implemented almost all vba.**, vb6.**, app.* objects and methods. Implements 64-bit compilation to generate EXE.

 The next thing VB6 needs is a relatively large project, such as RC6 .DLL some extended functions or test methods. IF YOU CAN DO A GITHUB PROJECT.
 You can directly upload, modify, and the main administrator is responsible for synthesis.
 If there is also a package manager NuGet. PIP.EXE ,NPM
 Everyone can directly check and install cZipArchive.cls, iSubclass_Mini.cls, stdPicEx.cls
 VB6 is in decline now, maybe our generation of old programmers can last for 10-30 years.
 It would be convenient to have a powerful online source management and questioning tool.

 In any case, many of my requests are still unreasonable, and very few people are used. I'm really sorry

 In fact, I also posted a lot of code bases on the forum, maybe some people feel that the code is too simple, and they are not qualified to be called "code bank",
 But there are parts that I also spent a lot of time summarizing and testing.
 We are all amateurs, selflessly contributing everyone's experience and technology to the convenience of others. Nor can you ask others to do too much.
 After all, only commercial companies like TwinBasic can achieve a large number of technological updates, but they cannot achieve the correct and reasonable requirements to meet and complete the development.

## #379 xiaoyao — Jun 6th, 2023, 04:18 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

How to drag a web page and move the entire form?
 I want to make a pure web page without a form border and put it inside the WEBVIEW
 Then move the web page and the whole interface moves with it, either by typing shortcuts or virtual minimize and maximize buttons in the web page.
 Do not know how to achieve

## #380 softv — Jun 6th, 2023, 07:09 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Dear Olaf,

 **A possible issue in the latest release** (29th May 2023) of RC6.dll

 I don't know whether the issue is because of a mistake at my end OR the following is indeed an issue in the latest release of RC6.dll

 **The issue is**:
 --
 When I use WV.Navigate to navigate to some external^^ site (not all, but some), the **WV.DocumentTitle** and **WV.DocumentURL** are not getting set.

 Example sites:
 <https://sureshramaswamy.org> https://sureshramaswamy.org
 <https://dandapani.org> https://dandapani.org

 (^^) as opposed to local html files in my own system.
 --

 1. When WV.DocumentTitle and WV.DocumentURL are not set for a particular site, if they are accessed (say, made to be printed via the statement Debug.Print "a-" & WV.DocumentTitle & "-a", "b-" & WV.DocumentURL & "-b", **it takes quite a few seconds** for the blank strings to be printed.

 2. If I navigate to 'https://google.com', the **"WV_DocumentComplete" event itself does not occur** (apart from the title and url not getting set). Strangely, this is the case with the previous RC6.dll also. No issues with URLs like "google.de", "google.in", etc.

 3. I found one or two other anomalies also while visiting some external sites (with the latest RC6.dll) but I felt they will get resolved automatically if the above issues are taken care itself. So, not mentioning about them, as of now.

 As of now, in my app (checked with your own WebView2 Demo project also), I have started using the previous RC6.dll itself and absolutely no issues with WV.DocumentTitle and WV.DocumentURL except for the solitary site "https://google.com"

 In case, the issues I have reported above are because of some mistake I am committing at my end only, then **kindly let me know what to do to rectify the same**.

 I take this opportunity to **thank you in tons**, once again, Olaf, for your contributions to the world society.

 Kind Regards.

## #381 Schmidt — Jun 6th, 2023, 08:54 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **softv** *[img: View Post]*

 **A possible issue in the latest release** (29th May 2023) of RC6.dll

 .
 As of now, in my app (checked with your own WebView2 Demo project also), I have started using the previous RC6.dll itself and absolutely no issues with WV.DocumentTitle and WV.DocumentURL except for the solitary site "https://google.com"

 In case, the issues I have reported above are because of some mistake...

No, can reproduce this here with your site-links.

 Background:
 In an attempt to circumvent the "suppression of added COM-Objects" (which "google.com" achieves, when it is loading) -
 I've changed the handling of the COM-Host-storage in 6.0.13 ...
 though without success as it seems...

 Have now restored the former vb_Host-behaviour in new version 6.0.14 (to what it was in 6.0.12).
 It's uploaded now (have not updated the page-labels to 6.0.14 though).

 Olaf

## #382 softv — Jun 7th, 2023, 02:32 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

... .. . In an attempt to circumvent the "suppression of added COM-Objects" (which "google.com" achieves, when it is loading) -
 I've changed the handling of the COM-Host-storage in 6.0.13 ...
 though without success as it seems...

 Have now restored the former vb_Host-behaviour in new version 6.0.14 (to what it was in 6.0.12). ... .. .
 Olaf

Thanks a ton, Olaf, for your mighty quick response. Yes, all back to normal now.

 Navigating to "google.com" alone does not populate the DocumentURL and DocumentTitle. But, from your reply I infer that it will be so for "google.com" alone. **Am I right?**

 Note:
 Using the latest RC6 (6.0.14), **WV.Navigate "https://google.com"** does trigger the 'WV_DocumentComplete' event but WV.DocumentURL and WV.DocumentTitle do not get populated. They are blank/empty.

 Kind Regards.

## #383 Schmidt — Jun 7th, 2023, 04:48 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **softv** *[img: View Post]*

Using the latest RC6 (6.0.14), **WV.Navigate "https://google.com"** does trigger the 'WV_DocumentComplete' event but WV.DocumentURL and WV.DocumentTitle do not get populated. They are blank/empty.

Yes, whilst "google.de" works, "google.in" works, etc. -
 the scripts Google loads (and runs), when loading from "google.com",
 are somehow crippling the COM-Object-Callback-mechanism of MS.

 And in turn, everything my wrapper offers communication-wise (.jsRun, .jsProp, etc.) is also crippled.
 (though only, when google.com was the navigation-target).

 As said, have started to find a way to circumvent this in 6.0.13,
 but as you found out - this affected other functionality with "formerly normal behaving URLs" (hence the back-paddling in 6.0.14).

 Think I'd like to wait a few months at this point,... perhaps MS comes up with a fix for that themselves -
 (or Google changes their offending script) ...if not, I'll restart my own attempts after that again...

 Olaf

## #384 xiaoyao — Jun 7th, 2023, 03:02 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

i have a full test：WebView2Demo.vbp，
 i install MicrosoftEdgeWebview2Setup.exe for only user1,not for all user.
 change to windows user2(local user):

 RC6_ 2022-0731
 ---------------------------
 Runtime error '430':
 Class does not support automation or expected interfaces
 ---------------------------
 RC6_ 2023-0529
 ---------------------------
 Err: 53, Error in FindFirstFile: ErrNum: 3, the system cannot find the specified path.
 Couldn't initialize WebView Binding

 this will check HKEY_CURRENT_USER ,HKEY_LOCAL_MACHINE,WOW6432Node,x64

Code:

```
Function GetEdgeWebViewPath() As String
   On Error Resume Next
   Dim WshShell As Object, WebView2_Version As String
   Dim RegPath As String, location As String
   Set WshShell = CreateObject("Wscript.Shell") 'HKEY_CURRENT_USER , HKEY_LOCAL_MACHINE
   RegPath = "HKEY_CURRENT_USER\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}\"
   location = WshShell.RegRead(RegPath & "location")
   If location = "" Then
        RegPath = Replace(RegPath, "\WOW6432Node", "")
        location = WshShell.RegRead(RegPath & "location")
        If location = "" Then
             RegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}\"
             location = WshShell.RegRead(RegPath & "location")
             If location = "" Then
                  RegPath = Replace(RegPath, "\WOW6432Node", "")
                  location = WshShell.RegRead(RegPath & "location")
             End If
        End If
   End If
   If location <> "" Then GetEdgeWebViewPath = location & "\" & WshShell.RegRead(RegPath & "PV")
End Function
```

## #385 Schmidt — Jun 7th, 2023, 03:23 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

i have a full test：WebView2Demo.vbp，
 **i install **MicrosoftEdgeWebview2Setup.exe** for only user1**, not for all user.

 **change to windows user2**(local user):

 ... no WebView2-instancing possible via RC6...

What's your question exactly?

 And why do you expect user2 to be able to create a WebView2-instance,
 when you did **not **ensure an installation of the WebView2-runtime for that specific user2?
 (or alternatively for **all users** of the machine)

 Olaf

## #386 xiaoyao — Jun 7th, 2023, 03:36 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Confirm that the EDGE Evergreen SDK was installed with non-administrator privileges last time, and only the installation data is found in the registry HKEY_CURRENT_USER (64-bit subkey).
 The path is: C:\Users\xiaoyao\AppData\Local\Microsoft\EdgeWebView\Application\114.0.1823.37
 EDGE information cannot be found in registry subkeys such as HKEY_LOCAL_MACHINE and 32-bit
 I switched to another Windows local administrator account and need to obtain directory permissions to access the previous user's directory: C:\Users\xiaoyao\AppData
 Then VB6 called WebView2Loader.dll loaded webview2 successfully, using fixed versions (old version) 109.0.1518.78.x86 and 110.0.1587.63.x86, tested that 3 versions of EdgeWebView SDK displayed web pages correctly.

 However, RC6 .dll specifying these 3 SDK directories to load webview2 failed.
 couldn't initialize WebView-Binding
 Install the Microsoft EdgeWebview2Setup .exe in administrator mode
 The path to the automatic installation is: C:\Program Files (x86)\Microsoft\EdgeWebView\Application
 ---------------------------
 The installation directory of EdgeWebView can also be found in the registry where the installer is uninstalled

 HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft EdgeWebView
 InstallLocation=C:\Users\xiaoyao\AppData\Local\Microsoft\EdgeWebView\Application
 Version=114.0.1823.37

 HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft EdgeWebView
 InstallLocation=C:\Program Files (x86)\Microsoft\EdgeWebView\Application
 Version=114.0.1823.41

## #387 xiaoyao — Jun 7th, 2023, 03:45 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The default double-click installation is non-administrator permissions, the path is: C:\Users\%user%\AppData\Local\Microsoft\EdgeWebView\Application\
 Right-click to install with administrator privileges, the path is: C:\Program Files (x86)\Microsoft\EdgeWebView\Application
 So you cannot directly specify the directory C:\Program Files (x86)\.
 You can try it this way:
 edgesdk="" or 2 other fixed (legacy) runtimes are fine
 If CreateCoreWebView2EnvironmentWithOptions(StrPtr(edgesdk), StrPtr(userdata), 0&

 After the user loads successfully, you can write down the current SDK installation directory. If the software user installs the software with non-administrator privileges and switches to another user to log in,
 Use the code to crack the permissions, or let the user click on the authorization to access the directory, so that there is no need to install it again.
 -----------------------------------

 If each user needs to install the SDK separately, it will occupy 600-800MB of hard disk space, and 2 users is 1.5GB of space
 Because we can't guarantee that the user will definitely install with administrator privileges, or it may just be a non-administrator account (the company implements permission control for users)
 If his account has administrator privileges, but does not have the right mouse button to click administrator privileges, the installation will not take effect.
 You can write your own installation script "runas" to force the installation of the runtime

## #388 Schmidt — Jun 7th, 2023, 04:12 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

Use the code to crack the permissions, or let the user click on the authorization to access the directory...

I'd suggest to other readers here, to not "crack any permissions"...

 The only things that needs to be done is, to run the WebView2-evergreen-installer properly "**as Admin**".
 (after downloading it from the MS-WebSite).

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

If each user needs to install the SDK separately, it will occupy 600-800MB of hard disk space...

And that large size is the reason, why it makes so much sense, to choose the "**for all Users**" option -
 on the given machine (once the installer was started "in Admin-mode").

 Olaf

## #389 xiaoyao — Jun 8th, 2023, 04:17 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Loading two WebVIEWs is slow, why?

Code:

```
Web1 InitComplete UsedTime：111.3823ms
Web2 InitComplete UsedTime：8124.4131ms
WEB1 NavigationCompleted Used：16458.6068ms
WEB2 NavigationCompleted Used：16462.0363ms
```

Code:

```
Option Explicit
Dim WithEvents Webview1 As RC6.cWebView2
Dim WithEvents Webview2 As RC6.cWebView2
Dim AA(10) As Currency, BB(10) As Currency, U1 As Currency, U2 As Currency
Dim Web1 As RC6.cWebView2
Private Sub Webview1_InitComplete()
        QueryPerformanceCounter AA(2)
        U1 = (AA(2) - AA(1)) / MsCount
        Debug.Print "Web1 InitComplete UsedTime：" & U1 & "ms"

        'Webview1.SyncSizeToHostWindow
        Webview1.Navigate "https://www.baidu.com?id=1"
End Sub
Private Sub Webview2_InitComplete()
        QueryPerformanceCounter BB(2)
        U2 = (BB(2) - BB(1)) / MsCount
        Debug.Print "Web2 InitComplete UsedTime：" & U2 & "ms"
        'Webview2.SyncSizeToHostWindow
        Webview2.Navigate "https://www.baidu.com?id=2"
End Sub
Private Sub Webview1_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
        QueryPerformanceCounter AA(3)
        U1 = (AA(3) - AA(1)) / MsCount
        Debug.Print "WEB1 NavigationCompleted Used：" & U1 & "ms"
End Sub

Private Sub Webview2_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
        QueryPerformanceCounter BB(3)
        U2 = (BB(3) - BB(1)) / MsCount
        Debug.Print "WEB2 NavigationCompleted Used：" & U2 & "ms"

End Sub

Private Sub Form_Activate()
If Me.Tag = "" Then
    Me.Tag = "a"
    QueryPerformanceCounter AA(1)
    QueryPerformanceCounter BB(1)
    Webview1.BindTo Picture1.hWnd, 0&, WebviewSDKPath, "Z:\TEMP"
    Webview2.BindTo Picture2.hWnd, 0&, WebviewSDKPath, "z:\temp2"

    ' Webview1.BindTo Picture1.hWnd, , WebviewSDKPath, "Z:\TEMP"
    ' Webview2.BindTo Picture2.hWnd, , WebviewSDKPath, "Z:\TEMP"
    'Webview1.Navigate "https://www.baidu.com"
    'Webview2.Navigate "https://www.baidu.com?id=2"
End If
End Sub

Private Sub Form_Load()
Set Webview1 = New_c.Webview2 'New cWebView2
Set Webview2 = New_c.Webview2 'New cWebView2
END SUB
```

IN my vb6 api webview2
 two webview use same userdata path:

Code:

```
2023/6/8 16:39:04 class1:web1
2023/6/8 16:39:04 class1:Web2
2023/6/8 16:39:04 class2:web1
2023/6/8 16:39:04>web1-InitComplete,116.7922ms
2023/6/8 16:39:04 class2:Web2
2023/6/8 16:39:04>Web2-InitComplete,  147.2914 ms
2023/6/8 16:39:04>web1-Completed usedtime 422.0084 ms
2023/6/8 16:39:04>Web2-Completed usedtime 465.4944ms
```

two webview use 2 userdata path:

Code:

```
----------vb6 api load webview2:
2023/6/8 18:09:04 class1:web1
2023/6/8 18:09:04 class1:Web2
2023/6/8 18:09:05 class2:web1
2023/6/8 18:09:05>web1-InitComplete, 437.4444 ms
2023/6/8 18:09:05>web1-Navigate  0.0443 ms
2023/6/8 18:09:05 class2:Web2
2023/6/8 18:09:05>Web2-InitComplete,  503.0666ms
2023/6/8 18:09:05>Web2-Navigate used 0.0215 ms
2023/6/8 18:09:05>web1-Completed,  827.0572ms
2023/6/8 18:09:05>Web2-Completed,  891.6135 ms

2023/6/8 18:09:14 RC6-Webview2 Test Start
Web1 InitComplete UsedTime：101.1542ms
Web1 Navigate UsedTime：3.1012ms
Web2 InitComplete UsedTime：454.1865ms
Web2 Navigate UsedTime：1.5263ms
WEB1 NavigationCompleted Used：537.3491ms
WEB2 NavigationCompleted Used：888.683ms
```

## #390 xiaoyao — Jun 8th, 2023, 04:32 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Webview1.Navigate "https://www.baidu.com?id=2",0
 Webview2.Navigate "https://www.baidu.com?id=2",0

 There is actually no need to wait here, if you want to wait, add a new function web.wait interface

 Web1 InitComplete UsedTime：45.2891ms
 Web2 InitComplete UsedTime：503.8269ms
 WEB1 NavigationCompleted Used：586.6115ms
 WEB2 NavigationCompleted Used：979.8215ms

## #391 Schmidt — Jun 8th, 2023, 09:42 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

Loading two WebVIEWs is slow, why?

Because you never read (or understand), <https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)&p=5606140&viewfull=1#post5606140> what was already explained to you in #362.

 The default-settings of the RC6-WebView2-wrapper will ensure synchronous calls... (to make it a bit more beginner-friendly).
 To change that to **a**sync-mode, one has to pass **0-WaitTimeouts** into the second parameters of the Bind- and Navigate-methods.

Code:

```
Option Explicit

Private WithEvents WV1 As cWebView2, WithEvents WV2 As cWebView2
Private WithEvents PB1 As PictureBox, WithEvents PB2 As PictureBox

Private Sub Form_Load()
  Set PB1 = Controls.Add("VB.PictureBox", "PB1"): PB1.Visible = True
  Set PB2 = Controls.Add("VB.PictureBox", "PB2"): PB2.Visible = True

  New_c.Timing True 'start high-res-timing
  Set WV1 = New_c.WebView2(PB1.hWnd, 0)
  Set WV2 = New_c.WebView2(PB2.hWnd, 0)
End Sub

Private Sub WV1_InitComplete()
  Debug.Print "WV1_InitComplete", New_c.Timing
  WV1.Navigate "https://sqlite.org/index.html", 0
End Sub
Private Sub WV2_InitComplete()
  Debug.Print "WV2_InitComplete", New_c.Timing
  WV2.Navigate "https://sqlite.org/index.html", 0
End Sub

Private Sub WV1_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
  Debug.Print "WV1_NavigationCompleted", New_c.Timing
End Sub
Private Sub WV2_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
  Debug.Print "WV2_NavigationCompleted", New_c.Timing
End Sub

Private Sub Form_Resize()
  If Not PB1 Is Nothing Then PB1.Move 0, 0, ScaleWidth, ScaleHeight \ 2
  If Not PB2 Is Nothing Then PB2.Move 0, ScaleHeight \ 2, ScaleWidth, ScaleHeight \ 2
  If Not WV1 Is Nothing Then WV1.SyncSizeToHostWindow
  If Not WV2 Is Nothing Then WV2.SyncSizeToHostWindow
End Sub
```

Timing-Results for the now async behaving 2 WV-instances (at the cost of using more Event-Handlers):

Code:

```
WV1_InitComplete             161.83msec
WV2_InitComplete             177.61msec
WV1_NavigationCompleted      270.31msec
WV2_NavigationCompleted      285.15msec
```

For the record:
 Performance-comparisons between basic-methods of the WebView2 are pointless (across wrapper-implementations) -
 in the end, it is the same MS-/Google-Chromium-libraries - which do all the heavy work under the covers.

 And as always:
 Please reduce your spamming of this thread -
 you now have your very own WebView2-thread here in the CodeBank
 (and I'd prefer when you do your "thinking out loud" over there).

 Olaf

## #392 xiaoyao — Jun 12th, 2023, 08:53 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

dim str as string
 str=jsProp("document.body.outerHTML")

 when webivew is busy
 or moretime
 it's run slowly,no result back

## #393 xiaoyao — Jun 12th, 2023, 07:15 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

The fourth question:
 (4) I need to place multiple WebViews on my Form or WidgetForm. I bind a WebView to a cwWidget through "WV.BindTo(W.Root.hWnd, 10)".
 The question is, when this cwWidget is removed, how to remove the bound WebView at the same time?

in vc++
 webviewController->get_CoreWebView2(&webviewWindow) it's com object
 set webviewController=nothing
 set webviewWindow=nothing
 so it will close edge browser in other process background

 use control array:

Code:

```
Web1pic(0).Visible = False
LOAD Web1pic(1)
Web1pic(1).Visible = True

Web1.BindTo(Web1pic(1).hWnd)
```

remove webview and close edge borwser:

Code:

```
Unload Web1pic(1)
```

Private Function FrmSubclassProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Form1, ByVal dwRefData As Long) As Long
 Select Case uMsg
 Case WM_COMMAND

 Case WM_DESTROY
 end sub
 if bindto(hosthwnd),hosthwnd have a msg WM_DESTROY,it will auto close edge webbrowser.

 if you bind 3 webview control(web1,web2,web3) only use one hostwnd( bind to same hwnd)
 you want to only close web2,mabe need call by rc6.dll for make new Method. or set web2=nothing

Code:

```

 Set Web1 = New_c.WebView2
 Set Web2 = New_c.WebView2
 Set Web3 = New_c.WebView2

Web1.BindTo(form2.hwnd)
Web2.BindTo(form2.hwnd)
Web3.BindTo(form2.hwnd)

movesize web1 to 0,0,400,400
move size web2 to 0,400,400,800
move size web3 to 0,800,400,1200
```

## #394 xiaoyao — Jun 12th, 2023, 08:50 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I spent a few hours re-learning the 10 pages that everyone responded to, and learned a lot of new things. I don't know if there are people like me?

 I feel that the learning cost of webview2 is very high, and the difficulty is too great. It is more than 10 times more difficult than webBrowser, and the similar functions are less than 5%. Microsoft is really disappointing, and VB.NET has not continued the ease of learning of VB6.

## #395 tmighty2 — Jun 14th, 2023, 05:55 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

When starting my app, I want to test if objects and libraries are setup correctly.

 To do that I have written some test functions like this:

Code:

```
Dim obj As Object

Set obj = CreateObject("LibraryRC6.cCairoContext")

If Not obj Is Nothing Then
    ' Library and class exist
    MsgBox "Library and class exist!"
Else
    ' Library or class does not exist
    MsgBox "Library or class does not exist!"
End If
```

However, that does not work, it throws the error "Active X can't create object".
 What would be the correct name to create this object or to check if the library can be used?

 I am not sure where to best post this question. If somebody tells me where to do that, I will try to move my post there.
 Thank you!

 ps: I don't want to use LoadLibrary as this does not check if the dll is registered correctly.

## #396 Arnoutdv — Jun 14th, 2023, 07:02 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Use error trapping

Code:

```
Option Explicit

Private Sub Form_Click()
  Debug.Print LibraryInstalled
End Sub

' Air code
Private Function LibraryInstalled() As Boolean
  Dim obj As Object

  On Error Resume Next

  Set obj = CreateObject("LibraryRC6.cCairoContext")
  If Err.Number = 0 Then
    If Not obj Is Nothing Then
      ' Library and class exist
      LibraryInstalled = True
    End If
  Else
    If Err.Number = 429 Then
      ' Library not installed
    Else
      ' Show an error
      MsgBox "Error: " & Err.Description & " (Number: " & Err.Number & ")"
    End If
  End If
End Function
```

## #397 Schmidt — Jun 14th, 2023, 09:34 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

ps: I don't want to use LoadLibrary as this does not check if the dll is registered correctly.

For the record, the correct "Library-Prefix" in a ProgID
 (which usually follows the scheme: "Lib.Class")
 ...is "RC6.cSomeClass"

 And the usual Class (which serves as a Main-Entry-Point), is "cConstructor".

 So, a CreateObject-based test, should ideally be done via ProgID: "RC6.cConstructor"

 That said, you can always work regfree with the RC6, when you:
 - ship the recent bunch of RC6(Base)-Dlls, all in a SubFolder of your App(zip)
 - and in addition - simply place a "helper-bas-module" in your App-Project.

 I've mentioned jpbro's work (related to this regfree-approach) already here in this thread:
 <https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)&p=5606658&viewfull=1#post5606658> https://www.vbforums.com/showthread....=1#post5606658

 HTH

 Olaf

## #398 BooksRUs — Jun 14th, 2023, 04:27 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hopefully, this is something simple. Trying to simply set the Background Color of the webview2 page. There is a property for it, but I can't seem to find where it is in the cWebView2 class.

 Thanks!

## #399 xiaoyao — Jun 14th, 2023, 07:09 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

Hopefully, this is something simple. Trying to simply set the Background Color of the webview2 page. There is a property for it, but I can't seem to find where it is in the cWebView2 class.

 Thanks!

it's use by css or js

 Supports WebDriver, so it also supports various scenarios such as crawlers and automated tests.
 Is it necessary to set the proxy IP in the startup parameters or something, and then you can support WebDriver?

## #400 xiaoyao — Jun 16th, 2023, 07:54 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Supports WebDriver, so it also supports various scenarios such as crawlers and automated tests.
 Is it necessary to set the proxy IP in the startup parameters or something, and then you can support WebDriver?

 Automate and test WebView2 apps with Microsoft Edge WebDriver
 <https://learn.microsoft.com/en-ca/microsoft-edge/webview2/how-to/webdriver> https://learn.microsoft.com/en-ca/mi...w-to/webdriver

 Step 4: Choosing whether Microsoft Edge WebDriver should launch your app or attach to it
 Decide whether to configure Selenium to drive WebView2 by using the "launch" or "attach" approach.

 The "launch" approach: In some scenarios, it's appropriate to let Microsoft Edge WebDriver handle launching your WebView2 app. Microsoft Edge WebDriver launches your WebView2 app and automatically attaches to the first available WebView2 instance that your app creates.

 The "attach" approach: In other scenarios, it's appropriate to attach Microsoft Edge WebDriver to a running WebView2 instance. You launch your app outside of Microsoft Edge WebDriver, and then attach Microsoft Edge WebDriver to a running WebView2 instance. This "attach" approach is for a WebView2 app that's not compatible with the "launch" approach.

 <https://learn.microsoft.com/en-ca/microsoft-edge/webview2/how-to/webdriver#step-4a-letting-microsoft-edge-webdriver-launch-your-webview2-app> https://learn.microsoft.com/en-ca/mi...r-webview2-app
 At this point, your app is running and its --remote-debugging-port command-line parameter has been set. Next, we'll attach Microsoft Edge WebDriver to the launched WebView2 app.

## #401 Schmidt — Jun 16th, 2023, 08:38 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

Hopefully, this is something simple. Trying to simply set the Background Color of the webview2 page.
 There is a property for it, but I can't seem to find where it is in the cWebView2 class.

There is not specific property for it on the WV2 ...

 The WV2-interface is generally much more "spartan", compared to the old IE-control -
 which IMO is a good thing - because the less you have to learn (about the MS-WebView2).

 Instead the focus now shifts from built-in COM-methods -
 to "easy reachable js-Methods- and -Properties, inside the Browser-engine"

 One has to be relatively proficient with JavaScript, to make good use of the WebView2
 (and that's a skill, many developers these days already have).

 Below is a code-snippet, which uses javascript-property-syntax - to influence the backcolor of the document.body.

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Me.Visible = True 'ensure visibility of this "Parent.hWnd"

  Set WV = New_c.WebView2(Me.hWnd) 'bind to the hWnd

      WV.NavigateToString "<p>Hello World</p>" 'create a simple document from string-input

      WV.jsProp("document.body.style.background") = "red" 'set the body.style.background-pop to red color
End Sub

Private Sub Form_Resize()
  If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub
```

HTH

 Olaf

## #402 softv — Jun 17th, 2023, 05:46 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Dear Olaf,

 Recently I noticed that if I have two applications (say, A and B) made by me (the 2nd application also invoking webview2 via RC6 WV only), then:

 1. If I start application A, the main interface opens normally, rendering the specified web page normally, and everything is working all right.

 2. But, **even as A is running, if I start B, then 'B' would not start**. In other words, the WebView2 rendering engine would not start for 'B'.

 So, I searched for a solution in net and landed here - <https://stackoverflow.com/questions/73779482/webview2-multiple-instances-of-same-application-need-to-launch-the-same-browser> https://stackoverflow.com/questions/...e-same-browser

 Seeing the suggestion in the abovementioned link, in 'B', I gave App.path for '**userDataFolder**' parameter while binding (as follows):

Code:

```
WV.BindTo(picWV.hWnd, , , App.Path)
```

When I did the above and started B, then B also would start simultaneously even while A is running. That made me happy.

 I noticed however that a directory named "**EBWebView**" got created in App.Path of 'B' application. The folders and files carried by that directory occupied 4Mb or so. That's also okay for me as long as 'B' started. But, that made me realise that I need to necessarily give 'App.path' during binding in 'A' application too because otherwise, if an application (say 'X') developed by some other developer was already running in user's system (assuming that 'X' did not have any folder specified for userDataFolder parameter), then if user starts my 'A' application simultaneously, then it would start only if I had specified a folder for userDataFolder in 'A'. But that would mean, 'EBWebView' would get created. "As of now", I don't foresee any problem, "as such", for me or the user because of that extra directory getting created but just wanted to know the following:
 --
 1. Is the suggestion (say 'S') given in the abovementioned stackoverflow page the ONLY way to start two or more WV-based applications simultaneously?

 2. If not, then, **what are the alternate ways? Will those ways be better than 'S'**? Perhaps they might avoid the creation of the 'EBWebView' but will they be better in terms of speed, etc.?
 Note: I do notice that with userDataFolder specified, during the first ever starting of 'A', it takes time for 'A' to start (because of the fresh creation of the 'EBWebView' directory, I believe). But, subsequent starts of 'A' are quite normal.
 --

 Well, just as I was about to complete drafting this message, I checked the size of 'EBWebView' in 'A' and it was 20MB!!! I checked in 'B' and it was 13MB! So, I think they are growing in size in every start of 'A'. If that is the case, then, it is a cause of concern. I will research more on this growing size when I get time and write back here. It seems like the **'Default' folder inside 'EBWebView' starts growing**, probably maintaining some session info, etc.

 All said and done, I believe that there should be a **proper way to bind** by which directories like 'EBWebView' need not get created.

 I just now quickly tried keeping different arbitrary values for SecondsToWaitForInitComplete in both 'A' and 'B' to see whether they would help as "**settings that differ per-instance**". But, in my quick testing, they did not seem to help.

 **Await a good solution**, if any.

 OR, in the first place, **in case I am wrong in my findings itself**, then **kindly let me know what mistake I did** and what is the corrective action to be taken on my side.

 I take this opportunity to once again **thank you in TONs**, Olaf, for all your wonderful free frameworks/codes/tools and prompt support. God Bless you! God Bless ALL!

 Kind Regards.

## #403 Schmidt — Jun 17th, 2023, 07:22 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **softv** *[img: View Post]*

1. If I start application A, the main interface opens normally, rendering the specified web page normally, and everything is working all right.

 2. But, **even as A is running, if I start B, then 'B' would not start**. In other words, the WebView2 rendering engine would not start for 'B'.

Yep, can affirm that...
 ...wasn't aware, how the concrete mechanism behind it works - but studied it a bit - and here's what I found:

 You have to provide a unique UserData-Directory for:
 - executables, which "**bytewise **differ" (that's probably tested by calculating HashSums in the startup-phase)

 Meaning, when you compile executable A.exe - and then:
 - just make a file-copy of the same A.exe - just changing the FileName to e.g. "Copy of A.exe"
 - this will **not **cause a conflict, when you start "A.exe" and "Copy of A.exe" in parallel
 (when the HashSum of two binaries is the same, they are allowed to **share** the same UserData-path - with the same "browser-settings").

 The "protection" kicks in, when you start executables which differ in their binary Hash-Sum (but point to the same "UserData"-Path).
 Then the behaviour is, that "the first one (on this UserData-path) wins".

 Hmm - how to solve that in the RC6...
 (regarding the "Default-UserData-Path" I apply, when you leave out the optional Param in the Bind-Method)

 Currently, I apply the result of: New_c.FSO.GetLocalAppDataPath ... as the "default-UserData-Path" (when the optional Param was not set)

 But I could (in the next release) change that, to e.g.: New_c.FSO.GetLocalAppDataPath & "\" & App.EXEName

 This way, we would achieve "isolation" between "different App.ExeNames" at least.

 If there's better ideas "forth-coming" here, I'd be happy to incorporate them -
 if not I'd proceed (for the next version), with what I've just proposed...

 Edit: forgot to comment about the "EBWebView"-folder below that UserData-path - #
 that's a folder, created by MS - and I wouldn't change its name - or "mess" with any of the content it holds from User-code...

 Olaf

## #404 softv — Jun 17th, 2023, 01:07 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Yep, can affirm that...
 ... .. .
 ... .. . But I could (in the next release) change that, to e.g.: New_c.FSO.GetLocalAppDataPath & "" & App.EXEName
 ... .. .

 Olaf

Thanks, as always, for your detailed and prompt reply, Olaf.

 That's (New_c.FSO.GetLocalAppDataPath & "" & App.EXEName) really unobtrusive too, while being unique. I myself saw the 'EBWebView' folder in my LocalAppData directory today only, after reading your message! *[img: Smilie]*. It was occupying 618MB! Don't know why it is designed to be growing though (by MS)!

 Kind Regards.

## #405 LeandroA — Jun 19th, 2023, 11:13 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello, what are the dependencies you have, one is CR6.dll but what else?

## #406 Schmidt — Jun 19th, 2023, 11:44 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **LeandroA** *[img: View Post]*

Hello, what are the dependencies you have, one is CR6.dll but what else?

Please always ship the whole content of RC6BaseDlls.zip (3.7MB zipped, currently).

 RC6.dll depends on DirectCOM.dll, cairo_sqlite.dll, as well as on WebView2Loader.dll (all 3 are "flat Dlls")

 I usually put them into a \Bin\ subdirectory below the App.Path (beside the executable) -
 and use one of the "Drop-in regfree-bas-modules" to enable regfree deployment of the whole App.

 Edit: Here's what a "minimum-version" of such a "regfree-DropIn-bas-module" looks like:
 (adjusted, for the libs being contained in the mentioned \Bin\ SubFolder)

Code:

```
Option Explicit

Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Function GetInstanceEx Lib "DirectCOM" (spFName As Long, spClassName As Long, Optional ByVal UseAlteredSearchPath As Boolean = True) As Object

Private Const DirectComDllRelPath = "\Bin\DirectCOM.dll"
Private Const RCDllRelPath = "\Bin\RC6.dll"

Public Property Get New_c() As cConstructor
  Static st_RC As cConstructor
  If Not st_RC Is Nothing Then Set New_c = st_RC: Exit Property

  If App.LogMode Then  'we run compiled - and try to ensure regfree instantiation from \Bin\
     On Error Resume Next
        LoadLibraryW StrPtr(App.Path & DirectComDllRelPath)
        Set st_RC = GetInstanceEx(StrPtr(App.Path & RCDllRelPath), StrPtr("cConstructor"))
        If st_RC Is Nothing Then MsgBox "Couldn't load regfree, will try with a registered version next..."
     On Error GoTo 0
  End If
  If st_RC Is Nothing Then Set st_RC = cGlobal.New_c 'fall back to loading a registered version
  Set New_c = st_RC
End Property

Public Property Get Cairo() As cCairo
  Static st_CR As cCairo
  If st_CR Is Nothing Then Set st_CR = New_c.Cairo
  Set Cairo = st_CR
End Property
```

Olaf

## #407 LeandroA — Jun 20th, 2023, 12:13 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

thank you very much, it was very helpful

## #408 xiaoyao — Jun 21st, 2023, 01:03 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

No spam here. 99% of the questions I ask or reply to are useful or have the correct answer. Maybe a little off topic at times.
 For example, someone asked about your previous WEBVIEW problem, why can’t EXCEL VBA be loaded, and why VB.NET will crash after loading? I replied and it was deleted by you. Because the form on VBA is a very special existence, it does not have a handle attribute, even if there is, there are still many problems, so if RC6.DLL itself can provide a window for displaying web pages, VBS can also display the form, VBA is also solved, and VB.NET will not crash.
 Didn't we solve that problem in the end?
 At this point, you may have said that you have strayed from the topic. Some things are completed in one sentence. If you don’t explain it clearly to you, you will get angry for no reason. I don’t know why. Maybe it’s a translation error?
 But you didn't expect me to be very troubled, why RC6.DLL webview always can't display the webpage, the loading is very slow, and the SDK directory cannot be found.
 For example, jsprop(**) often fails to return results.

## #409 Schmidt — Jun 21st, 2023, 01:52 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

No spam here. 99% of the questions I ask or reply to are useful or have the correct answer.

Nope, the opposite is the case... 99% of your postings:
 - are entirely out of context
 - in very bad english
 - are for the most part confusing or misleading
 - are not reacting to "questions, others were asking you"

 Please go "thinking out loud" elsewhere
 (but that's probably something, you will not react to either, although you were asked many times).

 Olaf

## #410 LeandroA — Jun 22nd, 2023, 08:46 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello, I want to provide some information or perhaps this has already been discussed, if you want to work with dpiAware = true, you can experience that the control does not initialize WV.BindTo() , it is important to take into account that if another application is instantiated using webview from RC6 with dpiAware = true and another application in dpiAware = false the last one will fail and vice versa, note if the IDE is executed, the manifested exe will fail until stopping the ide.

## #411 Schmidt — Jun 23rd, 2023, 12:18 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **LeandroA** *[img: View Post]*

Hello, I want to provide some information or perhaps this has already been discussed, if you want to work with dpiAware = true, you can experience that the control does not initialize WV.BindTo() , it is important to take into account that if another application is instantiated using webview from RC6 with dpiAware = true and another application in dpiAware = false the last one will fail and vice versa, note if the IDE is executed, the manifested exe will fail until stopping the ide.

That's similar to what softv brought up in #402 ...
 and should be fixed now in RC6 verion 6.0.15 - but I've not yet uploaded this new version...

 Olaf

## #412 xiaoyao — Jun 28th, 2023, 12:05 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

how to get selected item html?

Code:

```
var selection = window.getSelection();
var range = selection.getRangeAt(0);
var clonedRange = range.cloneRange();
var fragment = clonedRange.cloneContents();
var div = document.createElement('div');
div.appendChild(fragment);
var htmlString = div.innerHTML;
console.log(htmlString);
```

Code:

```
var htmlString = window.getSelection().getRangeAt(0).cloneContents().querySelector('body').innerHTML;

VM317:1 Uncaught TypeError: Cannot read properties of null (reading 'innerHTML')
    at <anonymous>:1:91
```

why?

## #413 Schmidt — Jun 28th, 2023, 01:36 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

how to get selected item html?

Code:

```
var selection = window.getSelection();
var range = selection.getRangeAt(0);
var clonedRange = range.cloneRange();
var fragment = clonedRange.cloneContents();
var div = document.createElement('div');
div.appendChild(fragment);
var htmlString = div.innerHTML;
console.log(htmlString);
```

Code:

```
var htmlString = window.getSelection().getRangeAt(0).cloneContents().querySelector('body').innerHTML;

VM317:1 Uncaught TypeError: Cannot read properties of null (reading 'innerHTML')
    at <anonymous>:1:91
```

why?

This is primarily a VB-Forum - and your question is related to "pure Javascript" -
 please ask it elsewhere and stop spamming this thread.

 Olaf

## #414 xiaoyao — Jun 28th, 2023, 02:10 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The browser control gets the selected text, HTM, or image, which is like a text box, and it is a very useful function to get part of the selected text. In addition to displaying web pages, web controls also need other JS-related functions.

 I have found the code, if it can be added to the DLL, it can make the webview more powerful

 In ie webbrowser, almost no JS functions are used, except that if you want to right-click to pop up the menu, you can actually use some undisclosed or special interfaces of WEBBROWSER to achieve it.
 However, edge webview2 requires a lot of JAVASCRIPT functions to be more convenient to use. If JS is disabled or JAVASCRIPT is not used, many functions cannot be realized.

 Perhaps it is possible to develop another RC6js.dll specifically to enhance JS capabilities, or add a general tools.js to enhance control capabilities, just like jquery.js exists.

 Sometimes I also want to get the text at the current mouse position (not selected) or just right clicked.
 Sometimes I want to get the coordinates of an element in the webpage ( left,top offset of the webview). There are many functions that are still necessary. If everyone can contribute some good JS functions to enhance the ability of edge, everyone's efforts will be great. Benefit 100-1000 other users.

 this is just a suggestion

## #415 Schmidt — Jun 28th, 2023, 04:25 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

Perhaps it is possible to develop another RC6js.dll specifically to enhance JS capabilities...

Why would anyone want to do that - when you can load such a "myUseful_JS_helper_functions.js" file
 into the WebView2-Context in just one line of code?

 This place is for the cWebView2-COM-Class of the RC6 - and javascript-problems should be mentioned only,
 when the COM-interface-methods for that (WV.jsProp, WV.jsRun, etc.) make problems.

 This Forum has a "javascript-SubForum" ... please post your "js-helpers-xiaoyao-style"
 in a separate thread over there.

 Olaf

## #416 tmighty2 — Jul 8th, 2023, 08:19 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Has anybody found a smart / easy way to use Olafs webview2 wrapper in Winform .NET?

 I wonder if I have to go interop, but I think so.

## #417 Schmidt — Jul 8th, 2023, 08:36 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

Has anybody found a smart / easy way to use Olafs webview2 wrapper in Winform .NET?

 I wonder if I have to go interop, but I think so.

I wonder why would you'd go .NET-Winforms (...is there a certain Control, not covered by Krools Suite)?

 Well, ... if you really have to - I'd use the .NET-based ("native") WebView2-wrapper instead of the RC6-based one.

 Olaf

## #418 Bob17 — Jul 10th, 2023, 09:11 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

That's similar to what softv brought up in #402 ...
 and should be fixed now in RC6 verion 6.0.15 - but I've not yet uploaded this new version...

 Olaf

Hi Olaf,
 do you plan to release the 6.0.15 with the fix for the DPI aware and DPI scaling issues?
 Thanks.

## #419 tmighty2 — Jul 14th, 2023, 09:07 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I would like to block the loading of images.

 I am using this after initialization:

 WV.AddWebResourceRequestedFilter "*", Filter_IMAGE

 However, this event is never called:

Code:

```
Private Sub WV_PermissionRequested(ByVal IsUserInitiated As Boolean, State As RC6.eWebView2PermissionState, ByVal URI As String, ByVal PermissionKind As RC6.eWebView2PermissionKind)

    Debug.Assert False

    If Has(URI, "jpg") Or Has(URI, "png") Then
        Debug.Assert False
    End If

End Sub
```

## #420 Schmidt — Jul 14th, 2023, 10:26 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

I would like to block the loading of images.

 I am using this after initialization:

 WV.AddWebResourceRequestedFilter "*", Filter_IMAGE

 However, this event is never called:

Code:

```
Private Sub WV_PermissionRequested(ByVal IsUserInitiated As Boolean, State As RC6.eWebView2PermissionState, ByVal URI As String, ByVal PermissionKind As
```

As explained in #373, the matching event is: WV_WebResourceRequested ...
 (and not WV_PermissionRequested, which is more for hardware-things like "GPS-sensor-access-permission, Microphon- or Camera-access-permissions")

 Olaf

## #421 Schmidt — Jul 23rd, 2023, 08:52 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Bob17** *[img: View Post]*

Hi Olaf,
 do you plan to release the 6.0.15 with the fix for the DPI aware and DPI scaling issues?

Not sure, whether the DPI-aware stuff is really related to the now "separated UserData-Folders per App-Exename" -
 but version 6.0.15 is now online.

 Olaf

## #422 xxdoc123 — Jul 25th, 2023, 07:15 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

i downlaod new current version: 6.0.15, last updated: 2023-07-23

 but i found can not run WV.jsProp. i change to 6.0.0.9 work ok.

 the other question

Code:

```
If Right$(App.Path, 1) <> "\" Then AppPath = AppPath & "\"

    If App.LogMode Then '
        On Error Resume Next
        LoadLibraryW StrPtr(AppPath & DirectComDllRelPath)
        Set New_c = GetInstanceEx(StrPtr(AppPath & RCDllRelPath), StrPtr("cConstructor"))
        'Set New_c = GetInstanceEx(StrPtr(AppPath & "Bin\RC6.dll"), StrPtr("cConstructor"), True) '
        If New_c Is Nothing Then MsgBox "Couldn't load regfree, will try with a registered version next..."
        On Error GoTo 0

    Else '
        Set New_c = New cConstructor
    End If
```

the same code if i chang RC6.dl 6.0.0.9 to 6.0.15 ，MsgBox "Couldn't load regfree, will try with a registered version next..."

 win7 32

 Another issue is that playing online videos often results in buffering. And the same browser plays smoothly。How to solve it？thanks

## #423 Schmidt — Jul 25th, 2023, 08:58 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xxdoc123** *[img: View Post]*

i downlaod new current version: 6.0.15, last updated: 2023-07-23

 but i found can not run WV.jsProp. i change to 6.0.0.9 work ok.

Please give an example...

>

>

 *[img: Quote]* Originally Posted by **xxdoc123** *[img: View Post]*

the other question

Code:

```
If Right$(App.Path, 1) <> "\" Then AppPath = AppPath & "\"

    If App.LogMode Then '
        On Error Resume Next
        LoadLibraryW StrPtr(AppPath & DirectComDllRelPath)
        Set New_c = GetInstanceEx(StrPtr(AppPath & RCDllRelPath), StrPtr("cConstructor"))
        'Set New_c = GetInstanceEx(StrPtr(AppPath & "Bin\RC6.dll"), StrPtr("cConstructor"), True) '
        If New_c Is Nothing Then MsgBox "Couldn't load regfree, will try with a registered version next..."
        On Error GoTo 0

    Else '
        Set New_c = New cConstructor
    End If
```

the same code if i chang RC6.dl 6.0.0.9 to 6.0.15 ，MsgBox "Couldn't load regfree, will try with a registered version next..."

The Dll-binaries within your Bin-Folder have to match with the (registered) RC6-version you last compiled your executable with...

>

>

 *[img: Quote]* Originally Posted by **xxdoc123** *[img: View Post]*

Another issue is that playing online videos often results in buffering.

That's not really a problem of the wrapper... (the Browser-engine is from Google and MS) -
 ...just guessing here, ...
 but perhaps a parallel installed MS-Edge/Chromium has different cache- or "virus-check-within-received stream-buffers" strategies at OS-level.

 Olaf

## #424 xxdoc123 — Jul 25th, 2023, 09:17 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

i test your demo WebView2Demo.zip

Code:

```
and this shows, that the WV.jsProp("...") also works in Property-Let-Mode (at the left-hand-side)
  btn1Caption = "Click Me..." 'change the Caption-String
  WV.jsProp("document.getElementById('btn1').innerHTML") = btn1Caption 'and assign it to the Browser-Element as the new Caption via WV.jsProp() = ...
  'just for fun, we can change the style of the btn1-Element to color='red' as well this way
  WV.jsProp("document.getElementById('btn1').style.color") = "red"
```

i found not change red.but old vision work fine

## #425 Schmidt — Jul 25th, 2023, 01:40 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xxdoc123** *[img: View Post]*

i test your demo WebView2Demo.zip

Code:

```
and this shows, that the WV.jsProp("...") also works in Property-Let-Mode (at the left-hand-side)
  btn1Caption = "Click Me..." 'change the Caption-String
  WV.jsProp("document.getElementById('btn1').innerHTML") = btn1Caption 'and assign it to the Browser-Element as the new Caption via WV.jsProp() = ...
  'just for fun, we can change the style of the btn1-Element to color='red' as well this way
  WV.jsProp("document.getElementById('btn1').style.color") = "red"
```

i found not change red.but old vision work fine

This problem came up already a few times here (first time, I think - in #185):
 <https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)&p=5580698&viewfull=1#post5580698> https://www.vbforums.com/showthread....=1#post5580698

 HTH

 Olaf

## #426 xxdoc123 — Jul 25th, 2023, 06:12 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

thanks. jsprop This function is very convenient to use。now must change my code

## #427 LeandroA — Jul 31st, 2023, 08:24 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

This was 3 steps back, the problem with the dpi has been solved, but jsProp no longer works as before, everything changed, I still don't know how to fix it. before this worked for me WV.jsProp("map.getCenter().lat()") now not anymore!
 Point #185 is not a solution, with WV.jsProp I could access a variable, if it worked fine, why did I change it?

## #428 Schmidt — Jul 31st, 2023, 09:36 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **LeandroA** *[img: View Post]*

jsProp no longer works as before, everything changed,...

Not "everything" - only jsProp shows a slightly different behaviour...

 And as explained in #185 - I had no choice in the matter ...
 (because newer Chromium-versions are now blocking eval() when used against some DOM-props).

 So, my workaround around this chromium-issue with eval() was, to now parse a "given property-string" by hand (step by step) -

 But note, that this workaround now only works on real property-sequences (and not "props with intermingled function-calls")...
 So the string still can be nested with "dots" in-between properties - but no parentheses (no func-calls) are allowed in the string.

 WV.jsProp("window.document.title") 'access on "pure" nested props still works

 WV.jsProp("map.getCenter().lat()") '...whilst this doesn't anymore, because it contains ()-identifyable, intermingled function-calls

>

>

 *[img: Quote]* Originally Posted by **LeandroA** *[img: View Post]*

 I still don't know how to fix it.

It's not that difficult, to write your own little helper-functions for communication with VB6,
 as e.g. (after WV.Init was successful)...

Code:

```
  With New_c.StringBuilder 'js-helpers

      .AddNL "function getMapCenter(){"
      .AddNL "    var c = map.getCenter()"
      .AddNL "    return [c.lat(), c.lon()].join(';')" 'returns two values (lat and lon) in a semicolon-separated string
      .AddNL "}"

      WV.AddScriptToExecuteOnDocumentCreated .ToString
  End With
```

Access then via WV.jsRun instead of WV.jsProp:

Code:

```
  Debug.Print WV.jsRun("getMapCenter")
```

Again, normally I try my very best, to not break existing functionality -
 but in this case it was the "increased security" of the chromium-engine which "broke things" (on some sites).

 **Edit**: What you can try, to simulate the "old behaviour" is, to run the javascript eval-function explicitely on your own like this:
 ... Debug.Print WV.jsRun("eval", "map.getCenter().lat()")
 ...meaning, that you could do a global replace of the term [ WV.jsProp( ] against [ WV.jsRun("eval", ]

 This might work on some sites (which don't enforce a "restricted eval()", like the sites delivered from from google-servers for example).

 Olaf

## #429 LeandroA — Jul 31st, 2023, 10:51 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thank you very much, this fixes it

 Private Function WV_Eval(ByVal Value As String) As String
 WV_Eval = WV.jsRun("eval", Value)
 End Function

 then replace all WV.jsProp by WV_Eval

## #430 xiaoyao — Jul 31st, 2023, 11:18 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

jsProp，This function has a lot of problems and bugs,

## #431 SearchingDataOnly — Aug 1st, 2023, 12:43 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xiaoyao** *[img: View Post]*

jsProp，This function has a lot of problems and bugs,

Your reply doesn't mean anything other than misleading others, and it's not what this forum needs.

 If you think jsProp has some problems, bring them up.
 If you think jsProp has bugs, report them.

 Based on my more than 10 years of experience with RC5/RC6, Olaf's code has very few bugs, and the quality of his code is at the top level even according to the standards of Commercial software.

## #432 xiaoyao — Aug 1st, 2023, 01:23 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

more times,jsProp("document.title"),result is empty,and other js

 WV_Eval replace to WV.jsProp maybe ok

## #433 xxdoc123 — Aug 8th, 2023, 09:19 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

execute WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();" sometimes the webpage is not executed successfully. So I will modify it to
 WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();"
 sleep 1000
 WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();"
 sleep 1000
 WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();"

 wo can tell me why？ and hao can do well？

## #434 Schmidt — Aug 8th, 2023, 03:33 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **xxdoc123** *[img: View Post]*

execute WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();" sometimes the webpage is not executed successfully. So I will modify it to
 WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();"
 sleep 1000
 WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();"
 sleep 1000
 WV.ExecuteScript "document.querySelector('#app > div > section > aside > span').click();"

 wo can tell me why？

WV.ExecuteScript is working asynchronously.
 And it should of course only be executed, when the Document was fully loaded
 (meaning, the call should be executed either from within WV_DocumentCompleted,
 or after that Event "came through").

 And since modern pages are able to enhance their DOM even after DocumentCompleted was signaled,
 it could very well be, that the span-element you want to click - simply "is not there yet".

 Please wrap the Selector-query in a small function, which is able to test for "existence" (of a certain, element-addressing query-string).

 E.g. you can add the 2 functions below via:
 WV.AddScriptToExecuteOnDocumentCreated _
 "function elementExists(sQuery){return document.querySelector(sQuery) ? true:false}" & vbCrLf & _
 "function clickElement(sQuery){document.querySelector(sQuery).click()}"

 And then test for existence via:
 Const sQuery As String = "#app > div > section > aside > span"
 If WV.jsRun("elementExists", sQuery) Then WV.jsRun("clickElement", sQuery) Else Debug.Print "Element doesn't exist yet..."

 HTH

 Olaf

## #435 tmighty2 — Sep 3rd, 2023, 04:52 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

There is no built-in method to store the entire contexts (html and images) of a webpage locally, is there?
 I thought that since the data is already on my disk when I ask WebView2 to navigate to a certain page, I could perhaps extract the local data.

## #436 xiaoyao — Sep 3rd, 2023, 05:42 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Please DON'T report bugs by a review, BUT HERE:
 - <https://github.com/vsDizzy/SaveAsMHT> https://github.com/vsDizzy/SaveAsMHT

 ######## Changelog #######

 <https://github.com/vsDizzy> https://github.com/vsDizzy

 CAN USE THIS?

## #437 Eduardo- — Sep 6th, 2023, 10:19 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello, I'm struggling to find the way to download a JSon file from an URL as string.
 I need to download it using the WebView2 because I'm already using it to log-in to the site and to download some pages, and I need to send the session Cookie to be able to download the JSon, so since I was not able to <https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)&p=5498591&viewfull=1#post5498591> get the Cookie with the method suggested here (because it appears that is no working now anymore, at least for me), I will have to get the JSon using the WebView2 itself and not by other download means.

 I tried to get it with WV.jsProp("document.body.outerText") and it kinda worked but if the JSon is too large it gets cut.
 Ideas?

## #438 Schmidt — Sep 7th, 2023, 06:31 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Eduardo-** *[img: View Post]*

Hello, I'm struggling to find the way to download a JSon file from an URL as string.
 I need to download it using the WebView2 because I'm already using it to log-in to the site and to download some pages, and I need to send the session Cookie to be able to download the JSon, ...

"in-Page" (in-document) - http-requests (aka "Ajax-Requests") are traditionally done via the XMLHttpRequest object:
 (which is quite similar in its usage, as the MS-WinHttp-COM-Object).
 <https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest> https://developer.mozilla.org/en-US/...XMLHttpRequest

 In other words - don't use the Webviews "Navigate"-methods to retrieve JSON from a certain WebAPI -
 use the XHR-Object instead "from the inside" of a loaded document (e.g. encapsulated in a little pre-added js-function).

>

>

 *[img: Quote]* Originally Posted by **Eduardo-** *[img: View Post]*

 ...so since I was not able to <https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)&p=5498591&viewfull=1#post5498591> get the Cookie with the method suggested here

Just tested this here - and it seems to work as described:

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Set WV = New_c.WebView2(hWnd, 0)
End Sub

Private Sub WV_InitComplete()
  WV.Navigate "https://vbForums.com", 0
End Sub

Private Sub WV_DocumentComplete()
  Debug.Print WV.jsProp("document.cookie")
End Sub

Private Sub Form_Resize()
  If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub
```

In case you're working against a "Google-URL" - these targets are known to torpedo the WebView2-COM-Object-support
 (the WV.AddObject-stuff ... which I need also on the inside of cWebView2, to interact with the VB-side).

 E.g. when you navigate to "https:/google.com" - then WV.jsProp and WV.jsRun won't work properly.

 Olaf

## #439 Eduardo- — Sep 7th, 2023, 06:55 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf. It is not Google, it is a private site.

 Debug.Print WV.jsProp("document.cookie") returns an empty string.

 Maybe that it is because it is an http site (not https)?

## #440 Schmidt — Sep 7th, 2023, 12:15 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Eduardo-** *[img: View Post]*

Hello Olaf. It is not Google, it is a private site.

 Debug.Print WV.jsProp("document.cookie") returns an empty string.

 Maybe that it is because it is an http site (not https)?

No, the reason in that case is, that document.cookie only returns "regular, **non**-HttpOnly-cookies"
 (http-only-cookies are hidden and non-accessible via javascript)...

 You can read about the different cookie-types here, for example:
 <https://dev.to/costamatheus97/battle-of-the-cookies-regular-cookies-vs-http-only-1n0a> https://dev.to/costamatheus97/battle...http-only-1n0a

 As for: "How to get to those http-only-cookies without using javascript..." -
 the WebView2 thankfully raises an Event we can use, where received resource-responses are reported
 (including a param, which transports a http-response-headers JSON-collection)...

 I've marked that Event in the following demo-code blue:

Code:

```
Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Set WV = New_c.WebView2(hWnd, 0)
End Sub

Private Sub WV_InitComplete()
  WV.Navigate "http://SomeDomain.com", 0
End Sub

Private Sub WV_DocumentComplete()
  Debug.Print WV.jsProp("document.cookie")
End Sub

Private Sub WV_WebResourceResponseReceived(ByVal ReqURI As String, ByVal ReqMethod As String, ByVal RespStatus As Long, ByVal RespReasonPhrase As String, ByVal RespHeaders As RC6.cCollection)
  If RespHeaders.Exists("Set-Cookie") Then Debug.Print RespHeaders("Set-Cookie")

'  Dim i As Long 'alternatively, one can enumerate all response-headers on the JSON-Collection
'  For i = 0 To RespHeaders.Count - 1
'    Debug.Print RespHeaders.KeyByIndex(i), RespHeaders.ItemByIndex(i)
'  Next
End Sub

Private Sub Form_Resize()
  If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub
```

Edit: One can also use the somwhat more modern "Fetch-API" instead of the XMLHttpRequest-Object:
 <https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API/Using_Fetch> https://developer.mozilla.org/en-US/...PI/Using_Fetch

 The http-cookie should be transported back to the server automatically in these "internal requests" -
 whereas "special tokens" have to be set manually in the clientside Headers of the request (in both cases)...

 But just test it out in your own small js-functions (added in WV_Initcomplete via AddScriptToExecuteOnDocumentCreated)

 HTH

 Olaf

## #441 Eduardo- — Sep 7th, 2023, 04:28 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

At the end I did it with a WinHttpRequest object that I already know how to handle and it worked.

 Thanks a lot!

## #442 VB6_Dev — Oct 7th, 2023, 10:28 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,

 Many thanks for developing this.

 The current **WebResourceResponseReceived** event only provides
 ReqURI, ReqMethod, RespStatus, RespReasonPhrase, RespHeaders

 Could you please implement a way of retrieving the **response body** (and, preferably, request headers & request body as well)?

## #443 VB6_Dev — Oct 7th, 2023, 10:30 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Reference:

 <https://learn.microsoft.com/en-us/dotnet/api/microsoft.web.webview2.core.corewebview2.webresourceresponsereceived> https://learn.microsoft.com/en-us/do...sponsereceived

## #444 tmighty2 — Oct 23rd, 2023, 04:10 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

*[img: Name:  audio1.png
Views: 2752
Size:  26.5 KB]*

 On a customer's computer I see this message when he / I click the microphone button in web.whatsapp.com.
 However, when he / I click on it, the result is not saved.
 It's not possible to record audio.

 Also,
 Private Sub WV_PermissionRequested
 is not fired.

 What could I be missing?
 The problem does not occur on my or (so far) any other customer computer.

## #445 tmighty2 — Oct 25th, 2023, 10:10 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf, could you offer a service that assures us "companies" to get support and a maintained product or the source code so that we could continue support it ourselves in case something happens to you?

## #446 Arnoutdv — Oct 25th, 2023, 10:42 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

WebView2 is not a product by Olaf, he just delivers a more easy programming interface using his vbRichClient 6 environment.

 <https://github.com/MicrosoftEdge/WebView2Browser> https://github.com/MicrosoftEdge/WebView2Browser
 <https://learn.microsoft.com/en-us/microsoft-edge/webview2/> https://learn.microsoft.com/en-us/mi...edge/webview2/

## #447 tmighty2 — Oct 25th, 2023, 11:03 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

That is exactely what I am talking about.

## #448 VanGoghGaming — Oct 25th, 2023, 12:28 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Outside this forum and a few other die-hard hobbyists it's pretty hard to still find a demand for VB6 apps, maybe that's why Olaf isn't selling his RC6 components.

 Looking at Arnoutdv's links I can see that this WebView2 stuff is pretty massive, so even if Olaf shared the source code I think most users would find it rather complicated to fix things themselves but at least it would open the gates for other experts to chime in and help out.

 I am wondering whether Olaf implemented WebView2 in VB6 or maybe it's all C++ in the RC6 DLL?

## #449 Schmidt — Oct 25th, 2023, 03:24 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

Private Sub WV_PermissionRequested
 is not fired.
 ...
 The problem does not occur on my or (so far) any other customer computer.

The problem also does not occur on my machine.

 Keep in mind, that "permission for accessing built-in hardware" (microphone, camera, GPS, etc.)
 has to take "more than one hurdle" (the "Privacy-related, BrowserCtl-internal permissions are only one side of the coin).

 There's also the "over-ruling" access-permission for such hardware on the outside (at OS-level).

 The OS allows, to define such permissions "App-specific" (for all installed "Store-Apps") -
 but also for either "all Desktop-Apps generally", or for "specific Process-Names of Desktop-Apps".
 (in a kind of WhiteList - probably stored in the registry somewhere).

 Remember, that your own "WebView2-using App" (your Executable) does not run under the flag of "MS-Edge" or something.
 (somehow inheriting the permissions of the MS-Browser-Executable)...
 No, it is an Executable in its own right (which by default is probably not in the "known list of white-listed Apps").

 So, try to find a way, to enable Microphone-usage at that outer level first for your App
 (e.g. in a Screen-Session, together with your one problem-customer).

 Some systems are quite "closed-off" due to over-eager Company-Domain-Admins -
 but that's not a problem the RC6-WebView2-Wrapper can solve...

 Olaf

## #450 Schmidt — Oct 25th, 2023, 03:33 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **VanGoghGaming** *[img: View Post]*

I am wondering whether Olaf implemented WebView2 in VB6 or maybe it's all C++ in the RC6 DLL?

All wrapper-classes the RC6.dll exposes, are VB6-implemented, since it's a normal VB6-AX-Dll-Project.
 (and that includes cWebView2).

 cWebView2 itself interacts with WebView2Loader.dll (which is part of the RC6.dll's "flat-Dll-dependencies") -
 and works through this relatively thin MS-provided layer, to talk (via COM-interfaces) -
 with the "Evergreen-WebView2-runtime" (about 300-500MB), which MS itself is updating regularly, once installed
 (on Win11, that "self-updating, chromium-based BrowserCtl-runtime" comes preinstalled on the system).

 Olaf

## #451 xiaoyao — Oct 25th, 2023, 07:16 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

dim web1 as cWebView2
 dim ptr1 as long-web1.webcontrolobjptr
 dim obj as object
 set obj=web1,WebView2control
 Can this object be made public? It is convenient for people in need to do more operations.

## #452 tmighty2 — Oct 26th, 2023, 06:22 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,

 can you release the source code of your components so that we can find a way to go on in case something happens to you?

## #453 Schmidt — Oct 27th, 2023, 11:18 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

...can you release the source code of your components so that we can find a way to go on in case something happens to you?

I'd say - let's talk about that, after "something happened to me"... *[img: Wink]*

 Besides, there's closed-source COMponents in (heavy) daily use in this community,
 which date nearly **20 **years back with their "last compile-date" as e.g. the unmaintained:
 - MS-Winsock.ocx
 - MS-FlexGrid, -HFlexGrid
 - MS-CommonControls
 - MS-ADODC

 And nobody demands from MS, that they open their sources for these controls.
 Instead those just work (and work and work) - because they are hardened -
 and for the few bugs they might still contain, there's "known workarounds".

 If you are that concerned, that a closed-source-component will not work anymore "next week" (as it was, at the last compile-date),
 then simply "choose another component" (as e.g. the .NET-wrapper for WebView2) -
 it really is that easy - and up to you to decide.

 I've mentioned the "conditions" for opening the RC6-sources already several times here in this forum...
 (the main-condition being: "a base for long-term-survival" - aka an OpenSourced, platform-independent compiler, which we do not have currently).

 Olaf

## #454 tmighty2 — Oct 29th, 2023, 11:39 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf,

 there is currently no replacement for VB6 for some developers.

 I am not releasing any .NET applications to the public because it means giving away your source code.
 I know 2 others who run big businesses using VB6 for the same reasons.

 I did convert my apps to .NET and decided against releasing them.

 People depend on my business.
 My softwares are certified and auditioned regularly as they are used in a medical branch.

 If you really mean it (regarding "after I am gone"), then can you implement a system that would make the source code accessible to me (and possibly others whose business depend on it)?

## #455 Schmidt — Oct 29th, 2023, 03:27 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

I am not releasing any .NET applications to the public because it means giving away your source code.

I hope, you notice the irony here...? *[img: Wink]*

 Besides, you can wrap up (only) the .NET WebView2-functionality in a VB6+COM consumable .NET-assembly
 quite easily... no need to rewrite "your whole App in .NET" (please read about .NET<->COM-interop).

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

If you really mean it (regarding "after I am gone"), then can you implement a system that would make the source code accessible to me (and possibly others whose business depend on it)?

There are already "measures in place" (which I tried to convey with my little "joke", which would otherwise be somewhat rude).

 Olaf

## #456 5501314zt — Nov 12th, 2023, 09:37 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

VB6.0,WebView2，Please advise on how to copy the PNG format webpage verification code into the picture box, boss.
 The website is as follows: **<http://xxpt.scxfks.com/study/captcha> http://xxpt.scxfks.com/study/captcha**

## #457 xiaoyao — Nov 12th, 2023, 12:42 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

A friend of mine in China has a long life span in his family, and many people spend more than 115 years.
 On average, half of them live to be over 100 years old.
 He packaged some of the source code into a CD and sold it to many people for a $50 tutorial fee.
 In fact, many technologies on our forum are many times more valuable than that.

 He is more than 70 years old this year, is a Taiwanese, and has not a single white hair.

 This really has a lot to do with heredity. And everyone's way of life.

 I wish everyone in here study vb6,vba,vb.net is happy, and everyone is happy to communicate.
 After net was abandoned by Microsoft, the number of VB programmers has been decreasing.
 In the past 10 years or so, we have probably contributed the most technology and the most powerful technology.

## #458 xiaoyao — Nov 12th, 2023, 12:53 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

@Schmidt You have contributed a lot of VB6 technology. Rc6, DLL. Is undoubtedly a very great product, if many people can learn to use it. It's really convenient, and it can also add a lot of functions to your software. A lot of development costs are reduced.

 It was supposed to be a paid software product for you to use for free, and we are very grateful to you.
 I was recently working on a mini version of the MySQL server, which was compressed to just over 2.8 megabytes.For the full version, it will be more than 50 megabytes normally.
 Mysqld. Exe, just start the process.
 And then set the super initial password. I studied it for two whole days. In fact, the cost is also very high. It can be interpreted as $500.
 There is one from 2003 or so. LIBMysql. DLL, less than 1000kb, it can connect to the MySQL server.It's been 20 years and it's still working. It's incredible.

 It's a little off topic.It's just that the research cost of some technologies is actually very high.You can choose open source or partial open source. It's all free.

 Many small functions are implemented with difficult technology and patented technology.
 We should respect the technology of original developers and give more encouragement.

## #459 xiaoyao — Nov 12th, 2023, 01:07 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **tmighty2** *[img: View Post]*

Olaf, could you offer a service that assures us "companies" to get support and a maintained product or the source code so that we could continue support it ourselves in case something happens to you?

To provide additional services, it must be a VIP who needs to pay extra. You can provide some give some source code or none.

 You can't unconditionally ask others to provide you with source code or provide you with more services.
 I've asked about this in Microsoft's MSDN support community, as well as in its open source project webviewload. DLL.
 It took two years to get back to me that they didn't have any plans to develop VB6 demo code.Although they can adopt or enhance functionality for some bugs or other technical issues.
 But they can also refuse 100% or never respond to your needs.
 Uch as continuing to offer vb7, VB. Net upgrading
 Or provide webview 2.ocx

## #460 saturnian — Dec 10th, 2023, 12:19 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf,
 I am working on a medical management software in France. For security reasons, reading social security cards and physician identification cards from a browser will soon no longer be possible through scripts within the web page, but through approved extensions added to the browser. Until now, I have been using cWebView2 to authenticate myself with these cards, but this will probably no longer be possible in the future.
 Apparently, WebView2 can now handle extensions.
 Is there a planned update of RC6 that will bring extension management to cWebView2? (Properties <https://learn.microsoft.com/en-us/dotnet/api/microsoft.web.webview2.core.corewebview2environmentoptions.arebrowserextensionsenabled?view=webview2-dotnet-1.0.2194-prerelease#microsoft-web-webview2-core-corewebview2environmentoptions-arebrowserextensionsenabled> AreBrowserExtensionsEnabled, etc…)
 Please keep me informed.
 Best regards.
 François

## #461 Schmidt — Dec 10th, 2023, 04:30 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

 Apparently, WebView2 can now handle extensions.
 ...

Ah well, ...finally.

 Will see what I can do in this regard over the holidays...

 Though, assuming you are aware of the AddObject-functionality (using WebView2 as a UI) -
 can't you just implement such "extended-stuff" in VB6-Classes and call it from the WebView?

 Olaf

## #462 saturnian — Dec 11th, 2023, 11:06 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thank you for your response, Olaf.

 I must admit that I am not familiar with the AddObject functionality!
 I will try to find examples in the posts of this thread while waiting for integration into cWebView2.

 Unless you have a small example?

 Best regards, François

## #463 Schmidt — Dec 11th, 2023, 12:02 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

 I must admit that I am not familiar with the AddObject functionality!

 Unless you have a small example?

It allows to put arbitrary Helper-COM-Objects (no matter, whether Private Classes or Public ones from an AX-Dll) -
 into the js-context... (under a certain name-string, which is then case-sensitive in js).

 Over such an added Object you can then - at any time:
 - "call into VB6-Code" (and thus, into "system-stuff, normally not reachable by a browser")

 example-class, named **cAuth **(in a normal VB-Form-Project):

Code:

```
Option Explicit

Public Function GetWinUserName() 'a Browser does not know the logon-name of the current Win-User
  GetWinUserName = Environ("username") 'so we help out with that here, via this little cAuth.Method
End Function
```

 TestForm-Code:

Code:

```
Option Explicit

Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Set WV = New_c.WebView2(hWnd, 0) '<-- the 0 ensures async-mode
End Sub

Private Sub Form_Resize()
  if Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub

Private Sub WV_InitComplete()
  WV.AddObject "Auth", New cAuth
  WV.NavigateToString "<div id='user_div'></div>", 0 '<-- the 0 ensures async-mode
End Sub

Private Sub WV_DocumentComplete()
  WV.ExecuteScript "document.querySelector('#user_div').textContent = 'LoggedOnUser: ' + Auth.GetWinUserName()"
End Sub
```

HTH

 Olaf

## #464 saturnian — Dec 11th, 2023, 02:41 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

"Thank you Olaf for this example.
 However, I don’t see how to switch a WebView2 property inside cWebView2, like AreBrowserExtensionsEnabled to True, through this method.

## #465 Schmidt — Dec 11th, 2023, 03:11 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

"Thank you Olaf for this example.
 However, I don’t see how to switch a WebView2 property inside cWebView2, like AreBrowserExtensionsEnabled to True, through this method.

As said, I'll try to enable this (not yet implemented) Property (and its "sidekicks") over the holidays in a new RC6-version, which enhances cWebView2.

 Why are extensions useful?
 Because they allow other devs to "break out" of the restricted Browser-environment with either js - or "other (compiling) languages" -
 to allow access to e.g. the FileSystem (along with anything else the Host-OS has to offer, which is not "reachable" via the built-in js-engine).

 The example in my prior post was just a suggestion... showing an alternative way how to tackle basically the same thing -
 (breaking out of the Browser, calling functionality of the underlying OS via VB and system-APIs) -
 entirely "on your own", in a language you know well, without waiting for "plugins" to circumvent "browser-limitations".

 Olaf

## #466 saturnian — Dec 11th, 2023, 03:44 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I understand better Olaf, Thank you.

 In fact, in my case, I am required to use the extensions approved by the administration to validate the reading of social security cards and physician identification in the browser context. It is also the back-end of the French national shared medical records management application that addresses the extension to authenticate patients and physicians before delivering health information.
 I will wait for the update of the RC6.

 Best regards.

 François

## #467 tmighty2 — Dec 25th, 2023, 06:27 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Edit: I think the problem occurs because I prevent new windows from popping up. I will report back.
 --------
 Still about not receiving the PermissionRequested event for web.whatsapp.com when trying to record an audio message on one customer's computer:
 My app is in the list of allowed applications.
 My app is able to access the microphone for example using waveInOpen and recording a wav.
 The system settings indicate that the admin takes care of system settings.
 What should I do now?
 Did I understand correctly that you offer to have a look at the customer's computer via Anydesk or similar?

## #468 julian1104 — Jan 8th, 2024, 10:21 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I am running a script as follows:

 dim script
 script = "var divsListAds = document.querySelectorAll('.mt-ListAds');" & vbCrLf & _
 "var hrefArray = [];" & vbCrLf & _
 "divsListAds.forEach(function(div) {" & vbCrLf & _
 " var links = div.querySelectorAll('a');" & vbCrLf & _
 " links.forEach(function(link) {" & vbCrLf & _
 " hrefArray.push(link.getAttribute('href'));" & vbCrLf & _
 " });" & vbCrLf & _
 "});" & vbCrLf & _
 "console.log(hrefArray);" & vbCrLf & _
 "window.chrome.webview.postMessage(JSON.stringify(hrefArray));"

 WV.ExecuteScript script

 In the console it shows me the data I need, the question is how can I pass that data to a variable or save that data in a text file, could someone help me?

## #469 Schmidt — Jan 10th, 2024, 06:00 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **julian1104** *[img: View Post]*

 WV.ExecuteScript script

 In the console it shows me the data I need, the question is how can I pass that data to a variable or save that data in a text file, could someone help me?

WV.ExecuteScript is not the right method to tackle such problems...

 A better approach is, to encapsulate your js-Codeblock in a function...

Code:

```
  With New_c.StringBuilder '
    .AddNL "function myFunc(){"
    .AddNL "    var divsListAds = document.querySelectorAll('.mt-ListAds')"
    .AddNL "    var hrefArray = []"
    .AddNL "    divsListAds.forEach(function(div){"
    .AddNL "        var links = div.querySelectorAll('a')"
    .AddNL "        links.forEach(function(link){"
    .AddNL "            hrefArray.push(link.getAttribute('href'))"
    .AddNL "        })"
    .AddNL "    })"
    .AddNL "    return JSON.stringify(hrefArray)"
    .AddNL "}"

    WV.AddScriptToExecuteOnDocumentCreated (.ToString) 'add the function to the WV-context from StringBuilder-Content
  End With
```

and then simply "calling it" later on (after Document-Load) via WV.jsRun()...

Code:

```
 Debug.Print WV.jsRun("myFunc") 'call the js-Function, printing the (stringified) JSON into the VB-Debug-Window
```

Olaf

## #470 julian1104 — Jan 10th, 2024, 11:53 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I greatly appreciate your message, I tried it but it doesn't print anything here:
 Debug.Print WV.jsRun("myFunc")

 I solved it this way but I don't consider it to be the most effective.

Code:

```
command1_click()
script = "var divsListAds = document.querySelectorAll('.mt-ListAds');" & vbCrLf & _
             "var hrefArray = [];" & vbCrLf & _
             "divsListAds.forEach(function(div) {" & vbCrLf & _
             "var enlaces = div.querySelectorAll('a');" & vbCrLf & _
             "enlaces.forEach(function(enlace) {" & vbCrLf & _
             "hrefArray.push(enlace.getAttribute('href'));" & vbCrLf & _
             "});" & vbCrLf & _
             "});" & vbCrLf & _
             "var hrefArrayString = hrefArray.join('\n');" & vbCrLf & _
             "navigator.clipboard.writeText(hrefArrayString).then(function() {" & vbCrLf & _
             "  window.chrome.webview.postMessage('Copiado al portapapeles');" & vbCrLf & _
             "});"

WV.ExecuteScript script
end sub
```

## #471 julian1104 — Jan 12th, 2024, 06:49 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I have the following code which works perfectly for me if I set breakpoints to validate the process, working perfectly this way, but if I remove the breakpoints and allow all the code to be processed, the system crashes. How could I solve it?

 This information extraction is executed from a loop for which it loads a website in for and loads the data to the text boxes

Code:

```
dim nextdata as boolean
private sub command1_click()
for a=1 to 2
WV.Navigate pagetoloader & "/" & a

Do
    DoEvents
Loop Until nextdata = True
        nextdata = False
next
end sub
```

Code:

```
Private Sub WV_DocumentComplete()

if nextdata=true then
script="xxxxxx"

WV.ExecuteScript script
text1=Here I stored the variable from script in a textbox which was copied directly from the clipboard
sleep 500

script2="xxxxxx"

WV.ExecuteScript script2
text2=Here I stored the variable from script in a textbox which was copied directly from the clipboard
sleep 500

script3="xxxxxx"

WV.ExecuteScript script3
text3=Here I stored the variable from script in a textbox which was copied directly from the clipboard
sleep 500

nextdata=true
end if
end sub
```

 The execution of the scripts are applied to each of the pages loaded in WV.Navigate pagetoloader & "/" & a

## #472 dday9 — Jan 12th, 2024, 09:35 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

By-passing captchas is not a topic that we allow on VBForums. Please do not discuss this topic any further.

## #473 julian1104 — Jan 16th, 2024, 03:39 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I have the next element is a website.

Code:

```
<div class="mt-ListPagination"><ul class="sui-MoleculePagination"><li class="sui-MoleculePagination-item"><a href="/segunda-mano/?st=2&amp;MakeIds%5B0%5D=4&amp;ModelIds%5B0%5D=0&amp;Versions%5B0%5D=&amp;pg=4" shape="circular" class="sui-AtomButton sui-AtomButton--primary sui-AtomButton--outline sui-AtomButton--center sui-AtomButton--small sui-AtomButton--link sui-AtomButton--empty sui-AtomButton--circular"><span class="sui-AtomButton-inner"><span class="sui-AtomButton-leftIcon"><span class="sui-AtomIcon sui-AtomIcon--small sui-AtomIcon--currentColor"><span class="sui-AtomIcon sui-AtomIcon--small sui-AtomIcon--currentColor" svgclass="mt-ListPagination-icon"><span><svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M7.293 11.707c0-.256.098-.512.293-.707l6-6L15 6.414l-5.293 5.293L15 17l-1.414 1.414-6-6a.997.997 0 0 1-.293-.707"></path></svg></span></span></span></span><span class="sui-AtomButton-content"></span></span></a></li><li class="sui-MoleculePagination-item"><a href="/segunda-mano/?st=2&amp;MakeIds[0]=4&amp;ModelIds[0]=0&amp;Versions[0]=" shape="circular" class="sui-AtomButton sui-AtomButton--primary sui-AtomButton--outline sui-AtomButton--center sui-AtomButton--small sui-AtomButton--link sui-AtomButton--circular"><span class="sui-AtomButton-inner"><span class="sui-AtomButton-content">1</span></span></a></li><li class="sui-MoleculePagination-divider">···</li><li class="sui-MoleculePagination-item"><a href="/segunda-mano/?st=2&amp;MakeIds%5B0%5D=4&amp;ModelIds%5B0%5D=0&amp;Versions%5B0%5D=&amp;pg=4" shape="circular" class="sui-AtomButton sui-AtomButton--primary sui-AtomButton--outline sui-AtomButton--center sui-AtomButton--small sui-AtomButton--link sui-AtomButton--circular"><span class="sui-AtomButton-inner"><span class="sui-AtomButton-content">4</span></span></a></li><li class="sui-MoleculePagination-item"><a href="/segunda-mano/?st=2&amp;MakeIds%5B0%5D=4&amp;ModelIds%5B0%5D=0&amp;Versions%5B0%5D=&amp;pg=5" shape="circular" class="sui-AtomButton sui-AtomButton--primary sui-AtomButton--solid sui-AtomButton--center sui-AtomButton--small sui-AtomButton--link sui-AtomButton--circular"><span class="sui-AtomButton-inner"><span class="sui-AtomButton-content">5</span></span></a></li><li class="sui-MoleculePagination-item"><a href="/segunda-mano/?st=2&amp;MakeIds%5B0%5D=4&amp;ModelIds%5B0%5D=0&amp;Versions%5B0%5D=&amp;pg=6" shape="circular" class="sui-AtomButton sui-AtomButton--primary sui-AtomButton--outline sui-AtomButton--center sui-AtomButton--small sui-AtomButton--link sui-AtomButton--circular"><span class="sui-AtomButton-inner"><span class="sui-AtomButton-content">6</span></span></a></li><li class="sui-MoleculePagination-divider">···</li><li class="sui-MoleculePagination-item"><a href="/segunda-mano/?st=2&amp;MakeIds%5B0%5D=4&amp;ModelIds%5B0%5D=0&amp;Versions%5B0%5D=&amp;pg=257" shape="circular" class="sui-AtomButton sui-AtomButton--primary sui-AtomButton--outline sui-AtomButton--center sui-AtomButton--small sui-AtomButton--link sui-AtomButton--circular"><span class="sui-AtomButton-inner"><span class="sui-AtomButton-content">257</span></span></a></li><li class="sui-MoleculePagination-item"><a href="/segunda-mano/?st=2&amp;MakeIds%5B0%5D=4&amp;ModelIds%5B0%5D=0&amp;Versions%5B0%5D=&amp;pg=6" shape="circular" class="sui-AtomButton sui-AtomButton--primary sui-AtomButton--outline sui-AtomButton--center sui-AtomButton--small sui-AtomButton--link sui-AtomButton--empty sui-AtomButton--circular"><span class="sui-AtomButton-inner"><span class="sui-AtomButton-content"></span><span class="sui-AtomButton-rightIcon"><span class="sui-AtomIcon sui-AtomIcon--small sui-AtomIcon--currentColor"><span class="sui-AtomIcon sui-AtomIcon--small sui-AtomIcon--currentColor" svgclass="mt-ListPagination-icon"><span><svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M16 11.707a.997.997 0 0 0-.293-.707l-6-6-1.414 1.414 5.293 5.293L8.293 17l1.414 1.414 6-6a.997.997 0 0 0 .293-.707"></path></svg></span></span></span></span></span></a></li></ul></div>
```

 I have tried clicking on the last element found within it (next page).

 but I don't have an answer to it, what am I doing wrong?

 example 1

Code:

```
Dim script As String

script = "var paginationItems = document.querySelectorAll('.mt-ListPagination .sui-MoleculePagination-item');" & vbCrLf & _
         "var ultimoItem = paginationItems[paginationItems.length - 1];" & vbCrLf & _
         "if (ultimoItem) {" & vbCrLf & _
         "    ultimoItem.click();" & vbCrLf & _
         "}"

WV.ExecuteScript script
```

 example 2

Code:

```
Dim script As String

script = "var ultimoItem = document.querySelector('.mt-ListPagination .sui-MoleculePagination-item:last-child');" & vbCrLf & _
         "if (ultimoItem) {" & vbCrLf & _
         "    ultimoItem.click();" & vbCrLf & _
         "}"

WV.ExecuteScript script
```

 example 3

Code:

```
Dim script As String

script = "var ultimoItem = document.querySelector('.mt-ListPagination .sui-MoleculePagination-item:last-child');" & vbCrLf & _
         "if (ultimoItem) {" & vbCrLf & _
         "    var eventoClick = document.createEvent('MouseEvents');" & vbCrLf & _
         "    eventoClick.initEvent('click', true, true);" & vbCrLf & _
         "    ultimoItem.dispatchEvent(eventoClick);" & vbCrLf & _
         "}"

WV.ExecuteScript script
```

 example 4

Code:

```
Dim script As String

script = "var ultimoItem = document.querySelector('.mt-ListPagination .sui-MoleculePagination-item:last-child');" & vbCrLf & _
         "if (ultimoItem) {" & vbCrLf & _
         "    ultimoItem.scrollIntoView();" & vbCrLf & _
         "    setTimeout(function() {" & vbCrLf & _
         "        var eventoClick = document.createEvent('MouseEvents');" & vbCrLf & _
         "        eventoClick.initEvent('click', true, true);" & vbCrLf & _
         "        ultimoItem.dispatchEvent(eventoClick);" & vbCrLf & _
         "    }, 500);" & vbCrLf & _
         "}"

WV.ExecuteScript script
```

## #474 saturnian — Feb 3rd, 2024, 01:30 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf,
 In the **BindTo **function, entering a command in **additionalBrowserArguments** (for example, "--lang=de") prevents the initialization of WebView2 with version 6.0.15 of RC6.
 What am I doing wrong?
 Thank you for your assistance.
 Kind regards.
 François

## #475 Schmidt — Feb 4th, 2024, 02:30 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

In the **BindTo **function, entering a command in **additionalBrowserArguments**
 (for example, "--lang=de") prevents the initialization of WebView2 with version 6.0.15 of RC6.

Just tested this here (on Win11) - and it works as it should...
 (on the **sole **instance of the WV-Ctl I've created)...

 IIRC, there was a problem sometime ago, where it was shown -
 that *every* "follow-up" WV2-instance in a given Process -
 needs to be started with the same commandline-params as the first created one...

 I consider that a problem MS has to fix (if it's still there).

 Olaf

## #476 saturnian — Feb 4th, 2024, 08:42 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

...
 IIRC, there was a problem sometime ago, where it was shown -
 that *every* "follow-up" WV2-instance in a given Process -
 needs to be started with the same commandline-params as the first created one...

 I consider that a problem MS has to fix (if it's still there).

 Olaf

Thank you Olaf,

 You are the best!

 It seems that all WV2s in a Windows session need to be initialized with the same parameters. I had two other applications running, different from the one I am developing, where a WV2 was initialized in French. It was then impossible to initialize in another language in the application I am developing. As soon as I closed these 2 applications, everything worked! This is a weird behavior from MS!

 Thanks again for your quick response.

 Best regards,
 François

## #477 Schmidt — Feb 4th, 2024, 11:11 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

It seems that all WV2s in a Windows session need to be initialized with the same parameters.

Ah, yes - that was the behaviour (which was discussed a few dozen entries back, here in this thread)...
 and a solution for this was (IIRC), that the behaviour can be avoided, as long as a different UserFolder was given
 (e.g. one Userfolder for french-users - and one for e.g. "--lang=de"-users on the same machine...).

 Olaf

## #478 saturnian — Feb 4th, 2024, 11:31 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Ah, yes - that was the behaviour (which was discussed a few dozen entries back, here in this thread)...
 and a solution for this was (IIRC), that the behaviour can be avoided, as long as a different UserFolder was given
 (e.g. one Userfolder for french-users - and one for e.g. "--lang=de"-users on the same machine...).

 Olaf

Thanks Olaf, I'll try that right away

## #479 Daniel Duta — Feb 12th, 2024, 08:04 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf,
 Is it possible to parse with WebView2 not a site but a local html page ? I would like to identify some controls within that html file and to change their values. For example to change a checkbox from true to false. Thank you.

## #480 saturnian — Mar 2nd, 2024, 01:09 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Ah well, ...finally.

 Will see what I can do in this regard over the holidays...

 Olaf

 Hello Olaf,
 Have you made progress on the new version of RC6, particularly improving the integration of WebView2 (Extension management...) ?
 Thank you for your response.

 Kind regards,
 François

## #481 JoseLiza — Apr 3rd, 2024, 12:24 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf.
 I am using this excellent option of using webview2 with vb6, I have managed to make functions to interact with the elements of some websites, but I found a problem that I have been struggling with for days, this is when simulating the click() event in an Input of type File, it simply does not open the file chooser dialog, when opening devtools the following alert can be displayed **"File chooser dialog can only be shown with a user activation"**, I understand that this is a security block, but I want to know if There is the option to bypass that restriction.

 Here I show the return message in the DevTools
 <https://www.vbforums.com/attachment.php?s=9b84d96bf559669a1dc45b4f80d160b8&attachmentid=191017> Attachment 191017

 Best regards.

## #482 Schmidt — Apr 3rd, 2024, 12:36 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **JoseLiza** *[img: View Post]*

 I have managed to make functions to interact with the elements of some websites, but I found a problem that I have been struggling with for days, this is when simulating the click() event in an Input of type File, it simply does not open the file chooser dialog, when opening devtools the following alert can be displayed **"File chooser dialog can only be shown with a user activation"**

 I understand that this is a security block...

Then don't fight it (within the Browser-Control)... *[img: Wink]*

 The nice thing about these "Dual-Apps" is, that you can always choose the side more suited for the task -
 so why not show a modal FileOpen-Dialog via VB6-code?

 Olaf

## #483 JoseLiza — Apr 3rd, 2024, 12:46 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Schmidt** *[img: View Post]*

Then don't fight it (within the Browser-Control)... *[img: Wink]*

 The nice thing about these "Dual-Apps" is, that you can always choose the side more suited for the task -
 so why not show a modal FileOpen-Dialog via VB6-code?

 Olaf

Using the vb6 FileOpen-Dialog would not allow assigning the path of the selected file in the file type input, since I am trying to upload a file to a website.
 If I'm wrong, I would appreciate if you could guide me on it.

## #484 cafeenman — Apr 3rd, 2024, 05:26 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The way I handle resizing is to set the form font to the control font and then use TextWidth with the text to determine the width and then size the control to whatever number it gives me back plus a little buffer.

 It only works on one-liners though - like labels and single-line textboxes. I've never tried it with multiline textboxes.

 I use it for all controls (that are one-lines) that I want to resize at run-time such as labels, check and optionboxes and single-line text boxes.

Code:

```
Private Function ResizeCheckBox(ByRef Check As Control, ByRef Form As Form, Optional ByRef Padding As Long = 0) As Long
Dim nWidth As Long

' Returns CheckBox width.
On Error GoTo errHandler

CopyFontFromControl Check, Form

nWidth = Form.TextWidth(Check.Caption)

nWidth = nWidth + CHECKBOX_PADDING + Padding

Check.Width = nWidth

ResizeCheckBox = nWidth

Exit Function

errHandler:
Dim nErrorNumber As Long
Dim nErrorHandlerResult As Long
Dim sError As String
Dim nErr As Long
Dim Parameters(1) As String

sError = Error
nErrorNumber = Err

Parameters(0) = "Check.Name = " & Check.NAME
Parameters(1) = "Check.Parent.Name = " & Check.Parent.NAME

nErrorHandlerResult = ErrorHandler(sError, nErrorNumber, ParameterString(Parameters), NAME & ".ResizeCheckBox(Private Function)")

Resume CleanUp

End Function

Public Function CopyFontFromControl(ByRef FromControl As Object, ByRef ToControl As Object) As Long

' Returns Error Code.
On Error GoTo errHandler

With ToControl.Font

  .Bold = FromControl.Font.Bold
  .Charset = FromControl.Font.Charset
  .Italic = FromControl.Font.Italic
  .NAME = FromControl.Font.NAME
  .Size = FromControl.Font.Size
  .Strikethrough = FromControl.Font.Strikethrough
  .Underline = FromControl.Font.Underline
  .Weight = FromControl.Font.Weight

End With

Exit Function

errHandler:
Dim nErrorNumber As Long
Dim nErrorHandlerResult As Long
Dim sError As String
Dim Parameters(1) As String

CopyFontFromControl = Err

nErrorNumber = Err
sError = Error

Parameters(0) = "FromControl.Name = " & FromControl.NAME
Parameters(1) = "ToControl.Name = " & ToControl.NAME

nErrorHandlerResult = ErrorHandler(sError, nErrorNumber, ParameterString(Parameters), NAME & ".CopyFontFromControl(Public Function)")

End Function
```

## #485 SomeYguy — Apr 24th, 2024, 08:34 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

(From dm)
 .....but the problem is that if a commercial software can only run in a Windows environment, its market demand and market share will be very low. Such softwares cannot form a solid software ecological chain, nor can they provide you with more resources.

Not quite true. Even in 2024 - some 4 years after this post - MS Windows still dominates the desktop, especially in productivity applications. And yes, there is still a smaller-than-in-the-past but re-growing market for desktop apps on Windows. Windows has always had - and likely always will - far more available software for most any conceivable desktop purpose then MAC OS and Linux combined. And I don't even want to go down the Linux mess rabbit hole - WAY too many "Distros", WAY to much incompatibility among them etc. so Linux is best left where it shines most - the server side.

 One can still build, sell and maintain good quality commercial 32-bit desktop apps with VB6. I've done it not so long ago and I plan to again. My users/customers can't tell/don't know/don't care what language their tools are written in.

 And now, with twinBasic up-and-coming, it's only going to get better!

## #486 SomeYguy — Apr 24th, 2024, 08:49 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **SearchingDataOnly** *[img: View Post]*

(From dm)
 .....but the problem is that if a commercial software can only run in a Windows environment, its market demand and market share will be very low. Such softwares cannot form a solid software ecological chain, nor can they provide you with more resources......
 \

Not quite true. Even in 2024 - some 3.5 years after this post, MS Windows still dominates the desktop, especially for productivity apps. There is still a smaller-than-before but re-growing market for commercial desktop applications. Windows has always had and likely always will have a much, much wider selection of available software than MAC OS & Linux combined. And the Linux desktop mess only makes it worse i.e. WAY to many "Distros", WAY too many incompatibilities between them, endless searches for needed dependencies, etc.

 One can still create, use and sell good quality commercial software with VB6. I've done it not so long ago and I plan to again. Some on this forum have also and still are.The core VB6 runtime will be supported until at least 2031 and likely beyond.

 My users/customers/groupies *[img: Wink]* can't tell/don't know/don't care what programming language their tools are written in. They just fire 'em up and away they go.

 And with twinBasic on the playing field, its just gonna get even better *[img: Smilie]*.

## #487 mark5544 — Apr 26th, 2024, 09:37 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

DOCUMENT.ACTIVEELEMENT ??

 Hi, Does anyone know how to use document.activeelement in webview2 please. Im struggling a bit, any code about how to set it up. Thanks very much

## #488 Tecnoimmobili — May 10th, 2024, 05:41 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

IT IS POSSIBLE TO EXTRAPOLATE TEXT FROM A WEB PAGE WITH WEBVIEW2. Thank you

## #489 tmighty2 — May 19th, 2024, 08:51 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf,

 can you please expose webview.defaultbackgroundcolor or show me how to access it in RC6 / WV2?

 Here is what I can do with WV in .NET and which I can not do with RC6 yet:

Code:

```
 Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim View As New WebView2()
        Me.Controls.Add(View)
        View.Dock = DockStyle.Fill
        Me.WindowState = FormWindowState.Maximized
        View.DefaultBackgroundColor = Color.Transparent
        View.BringToFront()
        Dim exeDirectory As String = Path.GetDirectoryName(Application.ExecutablePath)
        View.Source = New Uri(Path.Combine(exeDirectory, "cssanimation.html"))
        Me.FormBorderStyle = FormBorderStyle.None
        Me.AllowTransparency = True
        Me.BackColor = Color.Red
        Me.TransparencyKey = Color.Red
        Await View.EnsureCoreWebView2Async()
    End Sub
```

This makes the browser transparent.
 Or is there a way to set the defaultbackcolor any different way?

 I need to have a transparent browser. This will enable me to display the browser like a screen overlay.
 Currently I can not do this as this property is not exposed in the wrapper.
 Thank you.

 Also, I would like to point out that simply settings transparency like this...

Code:

```
<!DOCTYPE html>
<html>
<head>
    <style>
        body,
        html{
            background-color: transparent;
            user-select: none;
            margin: 0;
            padding: 0;
            border: 0;
            font-family: 'Nunito', sans-serif;
        }
        canvas{
            margin: 0;
            padding: 0;
            display: block; /*  ¯\_(?)_/¯  */
            touch-action: none;
            background-color: rgba(0, 0, 0, 0);
        }

        h1{
            position: absolute;
            top: 50%;
            left: 0;
            margin: auto;
            width: 100%;
            text-align: center;
        }
    </style>
</head>
```

does not do the job.
 The property that I am talking about causes the browser view to be really transparent: <https://learn.microsoft.com/en-us/dotnet/api/microsoft.web.webview2.core.corewebview2controller.defaultbackgroundcolor?view=webview2-dotnet-1.0.2478.35> https://learn.microsoft.com/en-us/do...et-1.0.2478.35

 Just in case you think that the property that I am talking about can be achieved by just defining it in html.

## #490 tmighty2 — May 21st, 2024, 02:46 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Yet another question:

 I want to subclass the wv2 window.
 I want to make it transparent, but I still the mouse move events, but it should also pass the window mesages on.
 What I did was to make me the hosting window layered and then I used subclassing.
 I received all the windows messages, and ignored the mouse clicks on the form, and I used the mouse move events as the browser required these in this case, but I also passed these on.
 However, this did not work as the hostHwnd is not what is required to subclass the messages that the browser receives I think.
 If anybody knows how to handle such a task, please let me know.

 Thank you.

## #491 saturnian — Aug 16th, 2024, 10:02 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf,
 In addition to extension management, would it be possible to integrate the asynchronous functions <https://learn.microsoft.com/en-us/dotnet/api/microsoft.web.webview2.core.corewebview2.printasync?view=webview2-dotnet-1.0.2592.51> PrintAsync and <https://learn.microsoft.com/en-us/dotnet/api/microsoft.web.webview2.core.corewebview2.printtopdfasync?view=webview2-dotnet-1.0.2592.51> PrintToPdfAsync into the cWebView2 component of the upcoming RC6 version?
 Simply integrating these functions (without defining CoreWebView2PrintSettings) would allow for the automation of certain printing tasks.

 Thank you for your response.

 Sincerely,
 François

## #492 misar — Sep 22nd, 2024, 10:23 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi Olaf,

 I have compiled your demo (WebView2Demo.zip) using the VB6 UI running on Win XP in order to test the exe on Win 10 and 11. It appears to be working correctly but in order to create the exe I had to rem out two of your event handlers which both generate the same error:
 Procedure declaration does not match description of event or procedure having the same name

 'Private Sub WV_AcceleratorKeyPressed(ByVal KeyState As eWebView2AccKeyState, ByVal IsExtendedKey As Long, ByVal WasKeyDown As Long, ByVal IsKeyReleased As Long, ByVal IsMenuKeyDown As Long, ByVal RepeatCount As Long, ByVal ScanCode As Long, IsHandled As Long)
 'Debug.Print "WV_AcceleratorKeyPressed"
 'End Sub

 'Private Sub WV_NavigationCompleted(ByVal IsSuccess As Long, ByVal WebErrorStatus As Long)
 'Debug.Print "WV_NavigationCompleted"
 'End Sub

 Do you know what is wrong here and also does your comment ['*** the rest of the EventHandlers below, are raised by the WebView-instance itself] mean that they are not needed anyway?

 Thanks,
 Mike

## #493 Schmidt — Sep 22nd, 2024, 10:36 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **misar** *[img: View Post]*

... I had to rem out two of your event handlers which both generate the same error:
 Procedure declaration does not match description of event or procedure having the same name

Have redefined the Interface since writing the demo... (e.g. in the two events - some of the Long-Param-Flags are now Booleans) -
 just re-select (after commenting out) these EventHandlers from the DropDown-Combobox again for correct Param-population...

>

>

 *[img: Quote]* Originally Posted by **misar** *[img: View Post]*

Do you know what is wrong here and also does your comment ['*** the rest of the EventHandlers below, are raised by the WebView-instance itself] mean that they are not needed anyway?

As with most Events - some of them are needed, some of them are left untouched...
 Which sub-set of them ends up being used, depends on the concrete scenario at hand...

 Olaf

## #494 misar — Sep 22nd, 2024, 11:40 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thanks Olaf.

 If anyone else has the same problem these are the corrected routines.

 Private Sub WV_AcceleratorKeyPressed(ByVal KeyState As RC6.eWebView2AccKeyState, ByVal IsExtendedKey As Boolean, ByVal WasKeyDown As Boolean, ByVal IsKeyReleased As Boolean, ByVal IsMenuKeyDown As Boolean, ByVal RepeatCount As Long, ByVal ScanCode As Long, IsHandled As Boolean)
 Debug.Print "WV_AcceleratorKeyPressed"
 End Sub

 Private Sub WV_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
 Debug.Print "WV_NavigationCompleted"
 End Sub

## #495 EasyOneX — Nov 14th, 2024, 09:35 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello.

 I have a quick question. In some old VB6 projects I used AutoComplete to connect a textbox with the history of IE. Is this also possible with WebView2 and if so how.

 Old VB6 Code:

 On the Form

Code:

```
Function EnableAutoComplete(hwnd As Long, dwFlags As Long) As Boolean
  On Error GoTo Err1

  SHAutoComplete hwnd, dwFlags
  EnableAutoComplete = True

Exit Function
Err1:
  EnableAutoComplete = False
End Function
```

Modul

Code:

```
Option Explicit

Public Declare Function SHAutoComplete Lib "shlwapi" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
'API constants
Public lngRet As Long

' // Currently (SHACF_FILESYSTEM | SHACF_URLALL)
Public Const SHACF_DEFAULT = &H0
' // This includes the File System as well as the rest of the shell (Desktop\My Computer\Control Panel\)
Public Const SHACF_FILESYSTEM = &H1
' // URLs in the User's History
Public Const SHACF_URLHISTORY = &H2
' // URLs in the User's Recently Used list.
Public Const SHACF_URLMRU = &H4
' // URLs in the User's Recently Used list.
Public Const SHACF_USETAB = &H8
' // Don't AutoComplete non-File System items.
Public Const SHACF_FILESYS_ONLY = &H10
Public Const SHACF_URLALL = (SHACF_URLHISTORY Or SHACF_URLMRU)

' // Ignore the registry default and force the feature on.
Public Const SHACF_AUTOSUGGEST_FORCE_ON = &H10000000
' // Ignore the registry default and force the feature off.
Public Const SHACF_AUTOSUGGEST_FORCE_OFF = &H20000000
' // Ignore the registry default and force the feature on.
' (Also know as AutoComplete)
Public Const SHACF_AUTOAPPEND_FORCE_ON = &H40000000
' // Ignore the registry default and force the feature off.
' (Also know as AutoComplete)
Public Const SHACF_AUTOAPPEND_FORCE_OFF = &H80000000
```

Many thanks in advance.

 Andreas

## #496 BooksRUs — Nov 21st, 2024, 12:23 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Olaf,

 Thanks again for the WebView2 control. This has really given new life to an old VB6 app (and developer *[img: Smilie]*

 I have a use for InPrivate mode or Incognito mode. I was wondering if there is an easy way to turn this on? My app logs into the same site multiple times with different credentials, and I have found that I have to make it logout/login for it to switch between them. I want to split this into 2 WVs and separately login to each.

## #497 misar — Nov 22nd, 2024, 04:12 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hello Olaf,
 Since I last posted I have used the WebView2 control very successfully to display web pages in my programs. I would now like to extend that to extracting the html content of a page (to a variable or save to file) after it has displayed. This is required with pages that use code to generate their content from a database (eg an Amazon S3 bucket). I assume it requires a script but could not work out how to do it from the examples in the demo. Is it possible and could you give me an example?
 Thanks

 Update:
 I already had the solution I need:

Code:

```
WV.jsProp("document.body.outerText")
```

but forgot that with a large amount of content the page needs a long time to complete!

## #498 BooksRUs — Nov 22nd, 2024, 06:52 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

misar,

 1. try using .jsProp("document.documentElement.innerHTML") - i do this all the time to pull out information i need from web pages. you just need to make sure that the document is "settled" and fully loaded.

 2. you actually replied to *my* post (just above) that asks about implementing "incognito" mode. I believe you should Reply to the very first post, so that Olaf is alerted.

 HTH!

 *[img: Smilie]*

## #499 misar — Nov 23rd, 2024, 01:32 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thanks BooksRUs,

 WV.jsProp("document.body.outerText") is sufficient for my present need (I just want to search the page content for specific text) but your suggestion is useful for future reference. It is easy to find lists of the DOM components but I have never come across clear descriptions distinguishing the specific data captured with different permutations. Hence I resort to trial and error to extract what I need!

## #500 misar — Nov 25th, 2024, 08:20 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I am working with a website (<https://bus.data.tfl.gov.uk/> https://bus.data.tfl.gov.uk/) which generates a file list from the contents of a large Amazon S3 data bucket. There is a substantial delay between the navigation completing and completion of the document that I want to capture. I thought that using Sub WV_DocumentComplete in Olaf's demo would deal with this but it fires immediately after Sub WV_NavigationCompleted. Is there a way to detect the true document completion? At present I work around this with a loop which checks the document size until it remains constant.

## #501 Eduardo- — Dec 2nd, 2024, 11:19 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **misar** *[img: View Post]*

I am working with a website (<https://bus.data.tfl.gov.uk/> https://bus.data.tfl.gov.uk/) which generates a file list from the contents of a large Amazon S3 data bucket. There is a substantial delay between the navigation completing and completion of the document that I want to capture. I thought that using Sub WV_DocumentComplete in Olaf's demo would deal with this but it fires immediately after Sub WV_NavigationCompleted. Is there a way to detect the true document completion? At present I work around this with a loop which checks the document size until it remains constant.

I tested and the event that keeps firing is WV_WebResourceResponseReceived.
 Maybe you could check when that event stops firing.

## #502 misar — Dec 3rd, 2024, 07:54 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Eduardo-** *[img: View Post]*

I tested and the event that keeps firing is WV_WebResourceResponseReceived.
 Maybe you could check when that event stops firing.

What are the correct parameters?

Code:

```
WV_WebResourceResponseReceived(ByVal ReqURI As String, ByVal ReqMethod As String, ByVal RespStatus As Long, ByVal RespReasonPhrase As String, ByVal RespHeaders As RC6.cCollection)
```

gives an error.

## #503 Eduardo- — Dec 3rd, 2024, 08:16 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

It must start with "Private Sub".
 The control adds the correct code automatically (I have no idea what you did).

## #504 misar — Dec 3rd, 2024, 08:31 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Eduardo-** *[img: View Post]*

It must start with "Private Sub".
 The control adds the correct code automatically (I have no idea what you did).

I started it with "Private Sub". The compiler throws an incorrect parameter error. I assumed this routine posted earlier in the thread had the correct parameters:

Code:

```
Private Sub WV_WebResourceResponseReceived(ByVal ReqURI As String, ByVal ReqMethod As String, ByVal RespStatus As Long, ByVal RespReasonPhrase As String, ByVal RespHeaders As RC6.cCollection)
On Error Resume Next
Text2.Text = Text2.Text & RespHeaders & vbNewLine
End Sub
```

## #505 Eduardo- — Dec 3rd, 2024, 09:04 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I have it working in one project, and it is the same:

Code:

```
Private Sub WV_WebResourceResponseReceived(ByVal ReqURI As String, ByVal ReqMethod As String, ByVal RespStatus As Long, ByVal RespReasonPhrase As String, ByVal RespHeaders As RC6.cCollection)
    ' (code)
End Sub
```

I'm using RC6 version 6.00.0014

## #506 misar — Dec 3rd, 2024, 09:20 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Eduardo-** *[img: View Post]*

I'm using RC6 version 6.00.0014

I'm using RC6 version 6.0.0.15 with the code from Olaf's demo. Apart from WV_WebResourceResponseReceived everything is compiling correctly.

 The version history says:
 version 6.0.0.15
 - cWebView2: now using *App-specific* UserData-Folders, in case the UserData-Param was left out in the Bind-Method
 - cSimplePSD: a bit of hardening, to properly cooperate with a "wider range of PSD-Files"
 - RC6Widgets got a new cwAlphaImg Widget

## #507 Eduardo- — Dec 3rd, 2024, 09:25 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

This example works for me:

## #508 misar — Dec 3rd, 2024, 09:50 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **Eduardo-** *[img: View Post]*

This example works for me:

Your example also works for me. The difference was the line Static c As Long
 - I added that to the routine and now my code also compiles and runs correctly.
 Many thanks for your help. *[img: Smilie]*

 Edit
 After looking at this more carefully I realised that the line "Static c As Long" is irrelevant to my code so I deleted it
 - and of course the routine still worked OK.
 I cannot understand what changed as previously I made several attempts to get apparently identical code to work. *[img: Confused]*

## #509 Eduardo- — Dec 3rd, 2024, 03:41 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **misar** *[img: View Post]*

I cannot understand what changed as previously I made several attempts to get apparently identical code to work. *[img: Confused]*

Sometimes there are mysteries... *[img: Alien Frog]*

 The event handler declaration that you pasted here was exactly the same.

## #510 BooksRUs — Dec 6th, 2024, 12:44 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

Olaf,

 Thanks again for the WebView2 control. This has really given new life to an old VB6 app (and developer *[img: Smilie]*

 I have a use for InPrivate mode or Incognito mode. I was wondering if there is an easy way to turn this on? My app logs into the same site multiple times with different credentials, and I have found that I have to make it logout/login for it to switch between them. I want to split this into 2 WVs and separately login to each.

Just bumping this... hopefully someone else has a need for this also. Before i posted before, and again just now, I tried passing different options for the parameter additionalBrowserArguments. I tried:

 "--enable-features=incognito"
 "--disable-features=msSmartScreenProtection"
 "--enable-features=incognito --disable-features=msSmartScreenProtection"

 None of which appear to do the job, even though no complaints from the New_c.WebView2 call.

 Any pointers would be great.

 Thanks!

## #511 jpbro — Dec 6th, 2024, 01:24 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

Just bumping this... hopefully someone else has a need for this also. Before i posted before, and again just now, I tried passing different options for the parameter additionalBrowserArguments. I tried:

 "--enable-features=incognito"
 "--disable-features=msSmartScreenProtection"
 "--enable-features=incognito --disable-features=msSmartScreenProtection"

 None of which appear to do the job, even though no complaints from the New_c.WebView2 call.

 Any pointers would be great.

 Thanks!

Have you tried using a different user data folder for each instance? You can then delete the "incognito" user folder on shutdown yourself. This approach is mentioned at the following links, but I haven't tried it myself:

 <https://stackoverflow.com/questions/75322492/how-webview2-runs-in-privacy-mode> https://stackoverflow.com/questions/...n-privacy-mode
 <https://github.com/MicrosoftEdge/WebView2Feedback/issues/779> https://github.com/MicrosoftEdge/Web...ack/issues/779

 I also found mention of an IsInPrivateModeEnabled option in the second thread, so maybe that's an avenue to explore?

## #512 saturnian — Dec 7th, 2024, 10:24 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jpbro** *[img: View Post]*

Have you tried using a different user data folder for each instance? You can then delete the "incognito" user folder on shutdown yourself. This approach is mentioned at the following links, but I haven't tried it myself:

 <https://stackoverflow.com/questions/75322492/how-webview2-runs-in-privacy-mode> https://stackoverflow.com/questions/...n-privacy-mode

this is how i work to create a private browsing mode !

Code:

```
Option Explicit
Private WithEvents WV As cWebView2
....
'when loading the container
Set WV = New_C.WebView2 'create the instance
Load picWV(1)
picWV(1).Visible = True

....

If mPrivate = True Then
      DataEdgePath = AddDirSep(Environ("TEMP")) & "MyAppName\" & CreateGUIDString(True) 'The navigation is in private mode and begins in a directory free of any data which will be destroyed at the end of navigation
Else
      DataEdgePath = fGetSpecialFolderLocation(CSIDL_COMMON_APPDATA) & "\"MyAppName\Data\" & Environ$("Username")
End If
MD DataEdgePath
.....

WV.BindTo picWV(1).hWnd, m_InitTimeOut, , DataEdgePath, TmpLang, mAllowSingleSignOnUsingOSPrimaryAccount

.....

'before unloading the container
Set WV = Nothing
Unload picWV(1)
If mPrivate = True Then KillFiles DataEdgePath, True

....

Private Sub MD(DirPath As String)
    If InStr(DirPath, "\") > 0 Then
    DirPath = AddDirSep(DirPath)
        On Error Resume Next
        Dim strTmp$(), f%, StrPath$
        strTmp$ = Split(DirPath, "\")
        For f% = 0 To UBound(strTmp()) - 1
            StrPath$ = StrPath$ & strTmp(f) & "\"
            MkDir StrPath
        Next f%
    Else
        MkDir DirPath
    End If
End Sub
Private Function AddDirSep(strPathName As String) As String
    If strPathName = "" Then Exit Function
    strPathName = RTrim$(strPathName)
    If Right$(strPathName, 1) <> "\" Then
            strPathName = strPathName & "\"
    End If
    AddDirSep = strPathName
End Function

........................................................................
    CreateGUIDString(True) return an unique GUID
    KillFiles delete all Folder,subFoldes and Files
    fGetSpecialFolderLocation retrieve specials Folders Paths
        ...see VBForums CodeBank...
```

## #513 talatoncu — Dec 12th, 2024, 05:33 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Dear Olaf,

 I have been using your product for a long time without problems.

 I have been loading Facebook page into WV and getting the body.innerhtml without any problem. Now, when I try to get the body.innerhtml, empty string is returned. I am waiting to get the html, but nothing returned. The code is as follows:

Code:

```
     WVerrorflag = False
     WV.Navigate "about:blank"
     WV.Navigate cururl$
     If WVerrorstatus <> 0 Or WV.DocumentURL = "chrome-error://chromewebdata/" Then
          Stop
     Else
          retcount = 0
          Do
               retcount = retcount + 1
               If retcount > 2 Then Stop
               curhtml$ = WV.jsProp("document.body.innerHTML")
               If curhtml$ <> "" Then Exit Do
               Wait 1
          Loop
     End If
```

Is there a way to get the html from Facebook page?

 Thank you for your kind help.

 Regards

## #514 talatoncu — Dec 14th, 2024, 08:52 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Dear Olaf,

 Is is possible to get the Content of the Response in

 WebResourceResponseReceived

 event?

## #515 BooksRUs — Dec 16th, 2024, 05:34 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **saturnian** *[img: View Post]*

this is how i work to create a private browsing mode !

*[img: Smilie]*

 Appreciate you guys putting me on the right path! It appears that using a different User Data Folder on each call *does* do the trick, so that the WV's run in separate folders. The unfortunate "side effect" is that it doesn't remember some of the user's default or previous selections, but the ability to run multiple separate instances far outweighs the inconvenience.

 Thanks!!

 Now, if I could just figure out how to keep/fix the ZoomFactor, I would be golden... if I hold control and roll the mouse to zoom in/out on a website, then go to a different place on the same site, the zoom resets.

## #516 jpbro — Dec 16th, 2024, 07:24 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

Appreciate you guys putting me on the right path! It appears that using a different User Data Folder on each call *does* do the trick, so that the WV's run in separate folders.

I was just looking at the WebView2 browser flags/params for another purpose and notice that there is a "--incognito" flag (maybe it is new?):

Code:

```
incognito 	Forces Incognito mode even if user data directory is specified by using the --user-data-dir flag.
```

Might be worth trying that to see if it keeps all of the user's existing settings but doesn't save/overwrite anything during/after the session.

## #517 jpbro — Dec 16th, 2024, 07:27 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BooksRUs** *[img: View Post]*

Now, if I could just figure out how to keep/fix the ZoomFactor, I would be golden... if I hold control and roll the mouse to zoom in/out on a website, then go to a different place on the same site, the zoom resets.

Regarding the zoom factor, it looks like others have reported the same problem:

 <https://github.com/MicrosoftEdge/WebView2Feedback/issues/3459> https://github.com/MicrosoftEdge/Web...ck/issues/3459

 I haven't had a chance to investigate, but that thread mentioned that there may be a solution in this thread:

 <https://github.com/MicrosoftEdge/WebView2Feedback/issues/2451> https://github.com/MicrosoftEdge/Web...ck/issues/2451

## #518 BooksRUs — Dec 17th, 2024, 09:26 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jpbro** *[img: View Post]*

I was just looking at the WebView2 browser flags/params for another purpose and notice that there is a "--incognito" flag (maybe it is new?):

Code:

```
incognito 	Forces Incognito mode even if user data directory is specified by using the --user-data-dir flag.
```

Might be worth trying that to see if it keeps all of the user's existing settings but doesn't save/overwrite anything during/after the session.

I think I had found the same documentation previously... I have tried passing each one of these as the additionalBrowserArguments parameter:

- "--disable-features=msSmartScreenProtection"
- "--enable-features=incognito"
- "--enable-features=incognito --disable-features=msSmartScreenProtection"
- "incognito"
- "--incognito"

 None of these worked for me. These are all from scouring the Internet trying to make this happen. Perhaps my formatting is off for passing in these parameters.

## #519 BooksRUs — Dec 17th, 2024, 10:13 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jpbro** *[img: View Post]*

Regarding the zoom factor, it looks like others have reported the same problem:

 <https://github.com/MicrosoftEdge/WebView2Feedback/issues/3459> https://github.com/MicrosoftEdge/Web...ck/issues/3459

 I haven't had a chance to investigate, but that thread mentioned that there may be a solution in this thread:

 <https://github.com/MicrosoftEdge/WebView2Feedback/issues/2451> https://github.com/MicrosoftEdge/Web...ck/issues/2451

*[img: Alien Frog]*

 Eureka!

 After digging into the 2nd link, it appears that I can now keep the same Zoom Factor between navigations (programmatically anyway):

Code:

```
dim ZoomFactor as double
...
' Before navigating to another page/site, save current ZF
ZoomFactor=WV.ZoomFactor
WV.ZoomFactor = ZoomFactor + .001

'navigate to site here
...

WV.ZoomFactor = ZoomFactor
```

I tried something similar before, but I think it needs a value that is *different* than what it currently is, in order to actually do anything. So bumping it by a little, navigating, and bumping it back by a little actually makes the WV respond.

 *[img: Smilie]*

## #520 BilalAhmed — Feb 13th, 2025, 10:06 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

can anybody please look into this thread <https://www.vbforums.com/showthread.php?910020-Iframes-with-sandbox-attributes-are-not-accessible-in-webview2-RC6-Help-please> https://www.vbforums.com/showthread....C6-Help-please and advise?

 Thankyou!

## #521 xiaoyao — Mar 5th, 2025, 03:17 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

how to use --disable-web-security??

 can't runjs to frame document.
 Uncaught DOMException: Failed to read a named property 'document' from 'Window': Blocked a frame with origin "https://**.com" from accessing a cross-origin frame.
 at <anonymous>:2:64
 ???? @ VM1604:2

 var iframe = document.querySelector("#pt_iframe");
 error here:
  var iframeDoc = iframe.contentDocument || iframe.contentWindow.document;

## #522 jpbro — Mar 5th, 2025, 07:17 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Did you try passing the switch to the additionalBrowserArguments parameter on the BindTo method?

 e.g.:

Code:

```
      MyWebView2.BindTo MyHwnd, , , , "--disable-web-security"
```

?

## #523 xiaoyao — Mar 5th, 2025, 11:22 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jpbro** *[img: View Post]*

Did you try passing the switch to the additionalBrowserArguments parameter on the BindTo method?

 e.g.:

 [/code]

 ?

thank you

 how to get response data by vb6?
 how to
 c# like " var response = await e.Request.GetResponseAsync();"
 how to get all url data ,like fiddler.exe or httpwatch.exe?

Code:

```
private async void InitializeWebView2()
{
    webView2 = new WebView2();
    webView2.Dock = DockStyle.Fill;
    this.Controls.Add(webView2);

    await webView2.EnsureCoreWebView2Async(null);
    webView2.CoreWebView2.WebResourceRequested += CoreWebView2_WebResourceRequested;
    webView2.CoreWebView2.Navigate("https://example.com");
}

private async void CoreWebView2_WebResourceRequested(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebResourceRequestedEventArgs e)
{
    // ??????
    if (e.Request.Uri.Contains("specific-url-to-monitor"))
    {
        // ??????
        var response = await e.Request.GetResponseAsync();
        var responseStream = await response.GetContentAsync();
        using (var reader = new System.IO.StreamReader(responseStream))
        {
            var responseContent = await reader.ReadToEndAsync();
            // ??????
            Console.WriteLine("Response content: " + responseContent);
        }
    }
}
```

## #524 5501314zt — Mar 30th, 2025, 04:13 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

<https://www.vbforums.com/attachment.php?s=9b84d96bf559669a1dc45b4f80d160b8&attachmentid=194487> Attachment 194487How do I trigger the ?start learning??“????”? button in VB6.0 version webview2?

## #525 xxdoc123 — Jul 18th, 2025, 09:37 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

now i have a question

Code:

```
Private Sub WV_DocumentComplete()

WV.jsRunAsync "LoadNextVideo"  'i used this js command to find video ,but Data is not available?

but if i set this command in commandbutton and click ?i wil get the data
Private Sub cmdCommand2_Click()

        infoLab.Caption = flag
        WV.jsRunAsync "LoadNextVideo"
End Sub
```

why ?

## #526 jpbro — Jul 18th, 2025, 09:48 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Perhaps the "data" loads asynchronously after the main document has finished loading? You might need to add a small helper script to look for a specific element that is injected into the DOM asynchronously and then call LoadNextVideo when it is found.

## #527 xxdoc123 — Jul 18th, 2025, 06:53 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jpbro** *[img: View Post]*

Perhaps the "data" loads asynchronously after the main document has finished loading? You might need to add a small helper script to look for a specific element that is injected into the DOM asynchronously and then call LoadNextVideo when it is found.

I'm not very proficient in js and don't know how to wait for data asynchronously

Code:

```
Do
            Wait 1000
            Debug.Print flag & "....."

            If WV.jsProp("document.querySelector('[class=""videotitle""]').innerText") = "ok" Then
                Debug.Print "we can se the title of the video?now run LoadNextVideo "

                WV.jsRunAsync "LoadNextVideo"  '

                Exit Do
            End If

        Loop
```

Still no data is available

## #528 BilalAhmed — Aug 3rd, 2025, 06:56 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

I know Olaf has already advised a way to pull http-only cookies in #440 like below

Code:

```
Private WithEvents WV As cWebView2

Private Sub Form_Load()
  Set WV = New_c.WebView2(hWnd, 0)
End Sub

Private Sub WV_InitComplete()
  WV.Navigate "http://SomeDomain.com", 0
End Sub

Private Sub WV_DocumentComplete()
  Debug.Print WV.jsProp("document.cookie")
End Sub

Private Sub WV_WebResourceResponseReceived(ByVal ReqURI As String, ByVal ReqMethod As String, ByVal RespStatus As Long, ByVal RespReasonPhrase As String, ByVal RespHeaders As RC6.cCollection)
  If RespHeaders.Exists("Set-Cookie") Then Debug.Print RespHeaders("Set-Cookie")

'  Dim i As Long 'alternatively, one can enumerate all response-headers on the JSON-Collection
'  For i = 0 To RespHeaders.Count - 1
'    Debug.Print RespHeaders.KeyByIndex(i), RespHeaders.ItemByIndex(i)
'  Next
End Sub

Private Sub Form_Resize()
  If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub
```

However, some modern websites (cannot share link) are using different ways to set secure cookies. Hence, they're not accessible the way Olaf recommended in #440. So I tried by directly accessing browser's cookies using following code:

Code:

```
WV.CallDevToolsProtocolMethod "Network.getAllCookies", "{}"
```

Problem is that CallDevToolsProtocolMethod does not return anything. So just cannot receive response which I need. So, is there a workaround I could receive response from CallDevToolsProtocolMethod method or an alternate but solid way to pull all the cookies (including http-only ones).

 Thankyou so much!

## #529 fafalone — Aug 4th, 2025, 12:02 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Well the underlying call does return something, through an async event completed interface, the problem with closed source wrappers is you're going to have trouble getting it since the author has vanished. Might be able to figure out a way to get a pointer to the underlying ICoreWebView2 interface to call it yourself

Code:

```
   Interface ICoreWebView2 Extends IUnknown
...
        Sub CallDevToolsProtocolMethod(ByVal methodName As LongPtr, ByVal parametersAsJson As LongPtr, ByVal handler As ICoreWebView2CallDevToolsProtocolMethodCompletedHandler)

    Interface ICoreWebView2CallDevToolsProtocolMethodCompletedHandler Extends IUnknown
        Sub Invoke(ByVal errorCode As Long /* HRESULT */, ByVal result As LongPtr /* LPWSTR */)
    End Interface
```

## #530 jpbro — Aug 4th, 2025, 10:52 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BilalAhmed** *[img: View Post]*

I
 However, some modern websites (cannot share link) are using different ways to set secure cookies. Hence, they're not accessible the way Olaf recommended in #440. So I tried by directly accessing browser's cookies using following code:

Code:

```
WV.CallDevToolsProtocolMethod "Network.getAllCookies", "{}"
```

Problem is that CallDevToolsProtocolMethod does not return anything. So just cannot receive response which I need. So, is there a workaround I could receive response from CallDevToolsProtocolMethod method or an alternate but solid way to pull all the cookies (including http-only ones).

 Thankyou so much!

If you add the following code to your project:

Code:

```
Private Sub WV_WebResourceResponseReceived(ByVal ReqURI As String, ByVal ReqMethod As String, ByVal RespStatus As Long, ByVal RespReasonPhrase As String, ByVal RespHeaders As RC6.cCollection)
   Debug.Print ReqURI, ReqMethod, RespStatus, RespReasonPhrase, RespHeaders.SerializeToJSONString
End Sub
```

And then call WV.CallDevToolsProtocolMethod "Network.getAllCookies", "{}" do you see any of the cookie data you are looking for in the Immediate window?

## #531 BilalAhmed — Aug 4th, 2025, 12:32 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Thank you jpbro, but if you could read my original message, you will notice that I have already tried this event. So to answer your question, unfortunately that doesnt work for me!

 Thank you so much!!!

## #532 jpbro — Aug 4th, 2025, 03:27 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Didn't know if you tried it with WV.CallDevToolsProtocolMethod "Network.getAllCookies", "{}", because I see the event fires multiple times and returns a lot of data when calling it against some example sites, but obviously I can't test with your site since you are unable to provide it. Hoped one or more of the firings would include your special Cookie(s).

## #533 BilalAhmed — Aug 4th, 2025, 04:13 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

My guess is that WV.CallDevToolsProtocolMethod "Network.getAllCookies", "{}" invokes JSAsyncResult event (only a guess, not really sure). But it doesnt bring anything to "WebResourceResponseReceived" and I think reason is obvious, WebResourceResponseReceived takes care only top level post/get/put/patch requests and listens to them. Whereas CallDevToolsProtocolMethod performs a low level operations and talks directly to the browser (Like selenium does). You're seeing so many calls being fired because several analytics softwares and some JS based sites keep refreshing data on the backend and, that is not possible without making a get/post etc.. requests which WebResourceResponseReceived event captures.

 So to 100% be sure, I carefully tested your code but, without luck!

 Thankyou so much!

## #534 fafalone — Aug 4th, 2025, 04:24 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The method RC6 is wrapping here is ICoreWebView2 -> Sub CallDevToolsProtocolMethod(ByVal methodName As LongPtr, ByVal parametersAsJson As LongPtr, ByVal handler As ICoreWebView2CallDevToolsProtocolMethodCompletedHandler)

 The results are returned to your implementation of ICoreWebView2CallDevToolsProtocolMethodCompletedHandler. RC6 does not look to have a wrapper for that, though I haven't looked too thoroughly-- see if a name like that exists in the events.

 Failing that, see if RC6 allows you to get a pointer to its internal ICoreWebView2 interface; if you can get such a pointer, you can use oleexp.tlb for the interfaces to invoke the command manually to your own implementation of the completion handler.

## #535 BilalAhmed — Aug 4th, 2025, 07:24 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Okay Faflone,

 You really sound logical but, unfortunately could not get a hold on anything you recommended. I also tried to see if WV exposes somthing by following 3 functions

Code:

```
Private Sub GetWebView2Pointer()
    Dim rawPtr As Long
    Dim WVPtr As Long
    Dim possibleVTablePtr As Long

    ' Step 1: Get pointer to the WV object
    WVPtr = ObjPtr(WV)

    ' Step 2: Read the first 4/8 bytes at the object's address
    ' This may be the vtable or a pointer to internal struct
    CopyMemory possibleVTablePtr, ByVal WVPtr, LenB(rawPtr)

    ' Print it for inspection
    Debug.Print "WV ObjPtr = &H" & Hex$(WVPtr)
    Debug.Print "First value at WV (possible vtable/internal struct) = &H" & Hex$(possibleVTablePtr)

    ' Try also to read 2nd level indirection if necessary:
    Dim possibleWebView2Ptr As Long
    CopyMemory possibleWebView2Ptr, ByVal possibleVTablePtr, LenB(rawPtr)
    Debug.Print "Value pointed by 1st value = &H" & Hex$(possibleWebView2Ptr)
End Sub
Private Sub ScanWVMemory()
    Dim baseAddr As Long
    baseAddr = ObjPtr(WV)

    Dim i As Long
    Dim Value As Long
    Debug.Print "Scanning memory starting at &H" & Hex$(baseAddr)

    For i = 0 To 60 Step 4
        CopyMemory Value, ByVal (baseAddr + i), 4
        Debug.Print "Offset +" & Hex$(i) & ": &H" & Hex$(Value)
    Next i
End Sub
Private Sub DumpWVStrings()
    Dim baseAddr As Long
    baseAddr = ObjPtr(WV)

    Dim addr As Long
    Dim tmpStr As String
    Dim i As Long

    For i = 0 To 200 Step 4
        CopyMemory addr, ByVal baseAddr + i, 4
        If addr > &H10000 And addr < &H7FFFFFFF Then
            On Error Resume Next
            tmpStr = Space$(200)
            CopyMemory ByVal StrPtr(tmpStr), ByVal addr, 200
            Debug.Print "Offset +" & Hex(i) & ": "; Left$(tmpStr, 100)
        End If
    Next i
End Sub
```

But this failed too. All it returned (especially "DumpWVStrings") just garbage *[img: Frown]*

 So probably Olaf may be able to help?

 Thankyou so much!

>

>

 *[img: Quote]* Originally Posted by **fafalone** *[img: View Post]*

The method RC6 is wrapping here is ICoreWebView2 -> Sub CallDevToolsProtocolMethod(ByVal methodName As LongPtr, ByVal parametersAsJson As LongPtr, ByVal handler As ICoreWebView2CallDevToolsProtocolMethodCompletedHandler)

 The results are returned to your implementation of ICoreWebView2CallDevToolsProtocolMethodCompletedHandler. RC6 does not look to have a wrapper for that, though I haven't looked too thoroughly-- see if a name like that exists in the events.

 Failing that, see if RC6 allows you to get a pointer to its internal ICoreWebView2 interface; if you can get such a pointer, you can use oleexp.tlb for the interfaces to invoke the command manually to your own implementation of the completion handler.

## #536 jpbro — Aug 4th, 2025, 09:59 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

If you're looking for an alternative, there's always twinBasic. It comes with an open source(-ish?) WebView2 implementation on top of a closed-source language/compiler that's mostly VB6 compatible. It's getting more compatible by the day and it only costs $35 USD/month, forever. Sure, it’s also a one-man show, but the "one man" is still really active compared to the other "one man", and this time it's different...If either you or he runs out of money, the whole thing just stops working completely *[img: Wink]*

## #537 fafalone — Aug 4th, 2025, 11:26 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Yeah it doesn't seem a method was included to get the pointer that you need, leaving only options not worth the trouble. Might want to use the twinBASIC WebView2 control; if you're not ready to switch to tB I've <https://github.com/fafalone/ucWebView2> wrapped it in an ActiveX control that works in VB6/VBA/etc. It comes with an event to receive the results of the call you're looking for by default with no need for modification. Switching to that will be easier than trying to dig out the pointer.

 ---

>

>
 If you're looking for an alternative, there's always twinBasic. It comes with an open source(-ish?) WebView2 implementation on top of a closed-source language/compiler that's mostly VB6 compatible. It's getting more compatible by the day and it only costs $35 USD/month, forever. Sure, it’s also a one-man show, but the "one man" is still really active compared to the other "one man", and this time it's different...If either you or he runs out of money, the whole thing just stops working completely

Maybe I missed the sarcasm but the idea either your code or the tB IDE/compiler just 'stops working completely' because either side runs out of money is absurdly false. There's a free community edition where the only limitation is a splash screen on x64 binaries, and no LLVM optimization (barely any of it implemented yet). As the "free" suggests, this costs $0 per month, not $35 per month. If you do get a premium subscription then stop paying, it just reverts to the community edition. Nothing stops working. And your binaries stay as is, there's no license checks in your code; an exe built with an active subscription will remain splashscreen free forever. If Wayne runs out of money, worst case is tB stays as it is now, already far ahead of VB6 in features and complete enough you can work around bugs and unimplemented parts. And that's if it happens soon before tB leaves Beta. Once it hits v1.0 the source goes into escrow to be released if abandoned.

## #538 BilalAhmed — Aug 5th, 2025, 05:45 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **fafalone** *[img: View Post]*

Yeah it doesn't seem a method was included to get the pointer that you need, leaving only options not worth the trouble. Might want to use the twinBASIC WebView2 control; if you're not ready to switch to tB I've <https://github.com/fafalone/ucWebView2> wrapped it in an ActiveX control that works in VB6/VBA/etc. It comes with an event to receive the results of the call you're looking for by default with no need for modification. Switching to that will be easier than trying to dig out the pointer.

 ---

 Maybe I missed the sarcasm but the idea either your code or the tB IDE/compiler just 'stops working completely' because either side runs out of money is absurdly false. There's a free community edition where the only limitation is a splash screen on x64 binaries, and no LLVM optimization (barely any of it implemented yet). As the "free" suggests, this costs $0 per month, not $35 per month. If you do get a premium subscription then stop paying, it just reverts to the community edition. Nothing stops working. And your binaries stay as is, there's no license checks in your code; an exe built with an active subscription will remain splashscreen free forever. If Wayne runs out of money, worst case is tB stays as it is now, already far ahead of VB6 in features and complete enough you can work around bugs and unimplemented parts. And that's if it happens soon before tB leaves Beta. Once it hits v1.0 the source goes into escrow to be released if abandoned.

Ahhh...

 you may be right your proposal Fafalone. But the problem under discussion is a smaller part of a large project (though very important). So it may take ages to port all the project and its dependencies in TB. Though I need the solution real quick but, I think at this moment, I should wait for Olaf's response with the hopes that he will come back soon.

 Thankyou so much!

## #539 fafalone — Aug 5th, 2025, 06:51 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

The control I posted is built in tB but for VB6. It compiles an ocx control you'd use in VB6 like any other.

 If you wanted to move the entire project to tB... The code should continue to work as is; tB is backwards compatible with VB6, there are no major rewrites like moving to another language. Ocx controls/dlls your project uses will work too. At most you might need to work around some bugs and a handful of missing capabilities; the only thing to actually rewrite is if you use VB internals hacks, which can be replaced with something simpler e.g. if you have a self-subclass think, tB supports AddressOf on class members so you'd just need a simple SetWindowSubclass call.

 Olaf hasn't been heard from in months now so the odds he'll return soon and quickly make a new RC6 version with the feature you want aren't great.

## #540 jpbro — Aug 6th, 2025, 11:41 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **fafalone** *[img: View Post]*

Maybe I missed the sarcasm but the idea either your code or the tB IDE/compiler just 'stops working completely' because either side runs out of money is absurdly false. There's a free community edition where the only limitation is a splash screen on x64 binaries, and no LLVM optimization (barely any of it implemented yet). As the "free" suggests, this costs $0 per month, not $35 per month. If you do get a premium subscription then stop paying, it just reverts to the community edition. Nothing stops working. And your binaries stay as is, there's no license checks in your code; an exe built with an active subscription will remain splashscreen free forever. If Wayne runs out of money, worst case is tB stays as it is now, already far ahead of VB6 in features and complete enough you can work around bugs and unimplemented parts. And that's if it happens soon before tB leaves Beta. Once it hits v1.0 the source goes into escrow to be released if abandoned.

You did miss some sarcasm *[img: Wink]*

 That said, unless I've misunderstood something, you can't build new stuff, fix old stuff, or release updates unless you keep paying the monthly fee indefinitely. A splash screen might be fine for hobby projects, but not much else.

 If Wayne runs out of money tomorrow, does the tB IDE keep running at your current license level forever, or are you out of luck?

>

>

 *[img: Quote]* Originally Posted by **fafalone** *[img: View Post]*

Once it hits v1.0 the source goes into escrow to be released if abandoned.

That's the promise, but there's no way to know if it will actually be kept, nor if it will be able to be independently verified should it ever come into play (again, unless I've missed something).

 For the record, I've had a yearly pro subscription to tB since the first or second month it was offered, and I've "bought Wayne a coffee" every month since that option was added too. I've been a financial supporter of tB from early on, and I genuinely believe it's the best shot we've ever had at true VB6 successor.

 But I also think it's a risky proposition for non-hobbyist development at this stage. And if we're being intellectually honest, the same arguments that have been/are being used against RC6 can and should be applied to tB as well.

## #541 BilalAhmed — Aug 6th, 2025, 02:50 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Okay, can you compile the ocx for me and upload somewhere to download? I dont mind paying $35 through credit card (probly one month subscription fee for TB).

 Thankyou so much!

## #542 fafalone — Aug 6th, 2025, 03:58 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **jpbro** *[img: View Post]*

You did miss some sarcasm *[img: Wink]*

 That said, unless I've misunderstood something, you can't build new stuff, fix old stuff, or release updates unless you keep paying the monthly fee indefinitely. A splash screen might be fine for hobby projects, but not much else.

 If Wayne runs out of money tomorrow, does the tB IDE keep running at your current license level forever, or are you out of luck?

 That's the promise, but there's no way to know if it will actually be kept, nor if it will be able to be independently verified should it ever come into play (again, unless I've missed something).

 For the record, I've had a yearly pro subscription to tB since the first or second month it was offered, and I've "bought Wayne a coffee" every month since that option was added too. I've been a financial supporter of tB from early on, and I genuinely believe it's the best shot we've ever had at true VB6 successor.

 But I also think it's a risky proposition for non-hobbyist development at this stage. And if we're being intellectually honest, the same arguments that have been/are being used against RC6 can and should be applied to tB as well.

You could ask about what happens to licenses if there's no more server... I'm not sure that it's even an issue right now; I'm sure there's some reasonable solution. But this isn't an issue unique to tB or to single developers... I'd argue that big corps are *more* likely to rug pull existing products and license servers in a way that forces customers off old products. There's certainly room to argue over details of escrow terms which haven't been discussed yet, but I'd be very surprised if that wasn't done at all, since it would be both an AH move and bad for tB's success, neither of which sound like something Wayne would do.

 Some of the same arguments apply but it's not intellectually honest to pretend there's not both material differences as well as an entirely different degree of how inconvenient alternatives are. RC6 is easy to not use. With tB, your only option that's not "throw out everything and start from scratch" is continuing with vb6, which is itself a risky proposition, given Microsoft already starting to end VBScript support. This topic we're on illustrates the other part: world of difference working around a missing feature in something like rc6 vs something as flexible as a programming language.

## #543 fafalone — Aug 6th, 2025, 04:02 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **BilalAhmed** *[img: View Post]*

Okay, can you compile the ocx for me and upload somewhere to download? I dont mind paying $35 through credit card (probly one month subscription fee for TB).

 Thankyou so much!

<https://github.com/fafalone/ucWebView2> https://github.com/fafalone/ucWebView2

 Open the ucWebView2 project in the .zip with <https://github.com/twinbasic/twinbasic/releases> twinBASIC (run as admin so it registers in HKLM where vb6 can see it) and click File->Build.

## #544 BilalAhmed — Aug 7th, 2025, 07:43 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **fafalone** *[img: View Post]*

<https://github.com/fafalone/ucWebView2> https://github.com/fafalone/ucWebView2

 Open the ucWebView2 project in the .zip with <https://github.com/twinbasic/twinbasic/releases> twinBASIC (run as admin so it registers in HKLM where vb6 can see it) and click File->Build.

fafalone, seems a good product (ucWebview). But I need a vb6 based sample project to understand it and learn good use of events/methods and properties. Can you please have me the link to it?

 Thankyou so much!

## #545 BilalAhmed — Aug 7th, 2025, 07:51 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Honestly, I could not even make it open google.com. If I do .navigate method on form load, it says "not ready yet". After an hour of struggle I found out that I cannot draw the control right on the form. Rather, I need to add a picturebox first and then draw the webview control over it.

 to summarize, there are so many precautions and things to consider to make it work. So a sample project which explains:

 1. Initialize the control and navigate to the page
 2. Get document content
 3. Access cookies (both secure and non secure)
 4. Make a devtoolprotocol call and receive response
 5. Manipulate DOM (fill forms, click buttons, set value to the text boxes, click ratio buttons etc)

 I may have more questions but, atleast I should be able to make it run at first.

 Thankyou so much!

## #546 BilalAhmed — Aug 7th, 2025, 08:13 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **fafalone** *[img: View Post]*

<https://github.com/fafalone/ucWebView2> https://github.com/fafalone/ucWebView2

 Open the ucWebView2 project in the .zip with <https://github.com/twinbasic/twinbasic/releases> twinBASIC (run as admin so it registers in HKLM where vb6 can see it) and click File->Build.

BUG: you cannot implement consume "webresourcerequested, navigationstarting" because it says "User defined type not defined"
 Similarly, JSAsyncResult, DownloadItemInterrupted, DownloadCompleted cannot be implemented in VB6 because it is restricted here.

 Apparently seem lots of bugs even though I am still unable to open my first page with this control.

 Thankyou so much!

## #547 jpbro — Aug 19th, 2025, 10:04 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **fafalone** *[img: View Post]*

You could ask about what happens to licenses if there's no more server... I'm not sure that it's even an issue right now;

I could, but I'd rather Wayne focus on the core than me bothering him about licensing server stuff...Again, I'm not trying to throw shade, but licensing *anything* from a one man show is a risk, even more so I think when it is a programming language compared to an add-on library. There will be nowhere to go with your codebase if tB goes belly up. That is a distinct difference compared to VB6+RC6 (or any other 3rd part closed source library) that will basically just keep working forever (albeit with whatever bugs are already baked into it, and limitations of the host language). Neither VB6 nor RC6 have any kind of active licensing system as a dependency, they just work (and will continue to just work essentially forever now with virtualization/emulation layers).

>

>

 *[img: Quote]* Originally Posted by **fafalone** *[img: View Post]*

I'm sure there's some reasonable solution.

I'm not so sure...if money runs out, will Wayne care to maintain the licensing server or release a patch to disable license checks? Maybe, maybe not.

>

>

 *[img: Quote]* Originally Posted by **fafalone** *[img: View Post]*

But this isn't an issue unique to tB or to single developers... I'd argue that big corps are *more* likely to rug pull existing products and license servers in a way that forces customers off old products.

Sure, but I'm talking about one man shows. Microsoft rug pulled as best they could, but many of us are still here using VB6 despite their attempts.

>

>

 *[img: Quote]* Originally Posted by **fafalone** *[img: View Post]*

There's certainly room to argue over details of escrow terms which haven't been discussed yet, but I'd be very surprised if that wasn't done at all, since it would be both an AH move and bad for tB's success, neither of which sound like something Wayne would do.

I agree that it would be a bad move, but at this stage it's just words. Not much different than Olaf's words about his own VB6 successor over the years.

>

>

 *[img: Quote]* Originally Posted by **fafalone** *[img: View Post]*

Some of the same arguments apply but it's not intellectually honest to pretend there's not both material differences as well as an entirely different degree of how inconvenient alternatives are. RC6 is easy to not use. With tB, your only option that's not "throw out everything and start from scratch" is continuing with vb6, which is itself a risky proposition, given Microsoft already starting to end VBScript support. This topic we're on illustrates the other part: world of difference working around a missing feature in something like rc6 vs something as flexible as a programming language.

RC6 and tB are totally different things IMO - RC6 is a library/add-on to VB6 that extends it in useful ways. tB is a new language/compiler that is a superset of VB6 that extends it in (probably more) useful ways. Difference is, that the more you adopt tB's "extensions", the more you get tied to tB forever, much in the same way that the more I used VB6, the more I got tied to VB6 forever. The difference is that VB6 is here and will be here for ever (at least in the manner I'm using it). tB is still at a very high risk stage, and IMO there's a significant chance that any work done to migrate code from VB6 to tB will be wasted. I'm not saying people shouldn't take that risk after evaluating all the options, but I do think it would be foolish to believe there isn't a significant risk.

## #548 carl039 — Aug 27th, 2025, 11:18 PM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

How could I use the following properties as detailed <https://learn.microsoft.com/en-us/dotnet/api/microsoft.web.webview2.core.corewebview2settings.ispasswordautosaveenabled?view=webview2-dotnet-1.0.1185.39> here

 IsPasswordAutosaveEnabled
 IsGeneralAutofillEnabled

## #549 Calcu — Sep 28th, 2025, 07:10 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi all, a little question i can't see asked/solved.

 I'm creating a web on a string.
 I'm using the webview2 like this:

Code:

```
 PicWv.Visible = True

    Set WV = New_c.WebView2
    If WV.BindTo(PicWv.hwnd) = 0 Then
    MsgBox "No puedo Inicializar el Webview", vbCritical, "Error!"
        Exit Sub
    End If
```

generating the web and launching this way:

 WV.NavigateToString laweb

 Everything is working fine, BUT i need to show another web (basically the same web with another data on it (diferent charts)), and then becomes the problem... the web is not reloading... i tried to modify the web content, destroying the WV, but nothing... always the same web until i close the form and launch again with the other data.... any idea about how to solve this ?

 Thanks!

## #550 Calcu — Sep 29th, 2025, 02:17 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Hi all, solved!

 the problem was:

 i was using this code in a button:

Code:

```
Set WV = New_c.WebView2

If WV.BindTo(PicWv.hwnd) = 0 Then
      MsgBox "No puedo Inicializar el Webview", vbCritical, "Error!"
      Exit Sub
End If
```

And everytime i pushed the button i was creating a new WV version...

 Just solved with:

Code:

```
    If WV Is Nothing Then
        Set WV = New_c.WebView2
        If WV.BindTo(PicWv.hwnd) = 0 Then
            MsgBox "No puedo Inicializar el Webview", vbCritical, "Error!"
            Exit Sub
        End If
    End If
```

Thanks!

## #551 talatoncu — Nov 3rd, 2025, 04:41 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Dear Schmidt,

 I got the following error when I try to reach the URL

 <http://www.muzayedeapp.com/muzayede-evleri-icin> http://www.muzayedeapp.com/muzayede-evleri-icin

 "This browser doesn't support the API's required to use the Firebase SDK"

 Is there a way to bypass this error?

 Thank you for your extremely valuable implementation and help.

 Talat Oncu

## #552 jpbro — Nov 3rd, 2025, 08:43 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **talatoncu** *[img: View Post]*

I got the following error when I try to reach the URL

 <http://www.muzayedeapp.com/muzayede-evleri-icin> http://www.muzayedeapp.com/muzayede-evleri-icin

 "This browser doesn't support the API's required to use the Firebase SDK"

Firebase is only supported on HTTPS websites. See: <https://stackoverflow.com/a/70750307> https://stackoverflow.com/a/70750307

>

>

 *[img: Quote]* Originally Posted by **talatoncu** *[img: View Post]*

Is there a way to bypass this error?

Change your URL from HTTP to HTTPS: <httpS://www.muzayedeapp.com/muzayede-evleri-icin> httpS://www.muzayedeapp.com/muzayede-evleri-icin and it should work.

## #553 jpbro — Nov 3rd, 2025, 11:03 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Should also mention that you should use HTTPS in general these days. Not just for security, but also because a lot of stuff will be broken when using HTTP.

## #554 wqweto — Dec 19th, 2025, 07:43 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

Here is how to reimplement missing **TitleChange** event as raised by the venerable built-in WebBrowser control.

Code:

```
Private WithEvents m_oWebView   As RC6.cWebView2

Private Sub m_oWebView_InitComplete()
    m_oWebView.AddScriptToExecuteOnDocumentCreated "var observer = new MutationObserver(function() { vbH().RaiseMessageEvent('title_change', document.title); });" & vbCrLf & _
        "document.addEventListener('DOMContentLoaded', function() { observer.observe(document.querySelector('title'), { childList: true }); vbH().RaiseMessageEvent('title_change', document.title); }); "
End Sub

Private Sub m_oWebView_JSMessage(ByVal sMsg As String, ByVal sMsgContent As String, oJSONContent As RC6.cCollection)
    If sMsg = "title_change" Then
        DebugPrint "New title is " & sMsgContent
    End If
End Sub
```

This should be possible to be exposed as a sorely missing separate event on **cWebView2** directly using **ICoreWebView2::add_DocumentTitleChanged** callback but until then using JS to setup MutationObserver on **title** element after DOMContentLoaded event is fired seems to be good enough workaround.

 cheers,
 </wqw>

## #555 saturnian — Dec 22nd, 2025, 08:50 AM

**Re: VB6 WebView2-Binding (Edge-Chromium)**

>

>

 *[img: Quote]* Originally Posted by **wqweto** *[img: View Post]*

Here is how to reimplement missing **TitleChange** event as raised by the venerable built-in WebBrowser control.

Code:

```
Private WithEvents m_oWebView   As RC6.cWebView2

Private Sub m_oWebView_InitComplete()
    m_oWebView.AddScriptToExecuteOnDocumentCreated "var observer = new MutationObserver(function() { vbH().RaiseMessageEvent('title_change', document.title); });" & vbCrLf & _
        "document.addEventListener('DOMContentLoaded', function() { observer.observe(document.querySelector('title'), { childList: true }); vbH().RaiseMessageEvent('title_change', document.title); }); "
End Sub

Private Sub m_oWebView_JSMessage(ByVal sMsg As String, ByVal sMsgContent As String, oJSONContent As RC6.cCollection)
    If sMsg = "title_change" Then
        DebugPrint "New title is " & sMsgContent
    End If
End Sub
```

This should be possible to be exposed as a sorely missing separate event on **cWebView2** directly using **ICoreWebView2::add_DocumentTitleChanged** callback but until then using JS to setup MutationObserver on **title** element after DOMContentLoaded event is fired seems to be good enough workaround.

 cheers,
 </wqw>

Thanks Wqweto for this tip!
 Happy holidays!
