---
id: how-to-work-with-groupdocs-assembly-for-net-standard-2-0
url: assembly/net/how-to-work-with-groupdocs.assembly-for-net-standard-2.0/
title: How to Work with GroupDocs.Assembly for .NET Standard 2.0
weight: 7
description: "Learn what additional dependencies you may need to install to work with GroupDocs.Assembly for .NET Standard 2.0 in your project, depending on your OS."
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
GroupDocs.Assembly is built on top of several Aspose products. Some of the products use [SkiaSharp](https://www.nuget.org/packages/SkiaSharp/) to deal with graphics-related tasks within .NET Standard 2.0, whereas the others prefer [System.Drawing.Common](https://www.nuget.org/packages/System.Drawing.Common/). That is why, GroupDocs.Assembly depends on the both packages.

A recommended way to deal with these and other GroupDocs.Assembly for .NET Standard 2.0 dependencies is installing of [GroupDocs.Assembly NuGet package](https://www.nuget.org/packages/GroupDocs.Assembly/). Otherwise, you can install all dependencies listed at [the NuGet package page](https://www.nuget.org/packages/GroupDocs.Assembly/) manually.

Depending on OS you work on, you may need to install some additional dependencies as explained further.

## Use GroupDocs.Assembly for .NET Standard 2.0 in Windows

No additional dependencies arise.

## Use GroupDocs.Assembly for .NET Standard 2.0 in Linux

In addition to GroupDocs.Assembly NuGet package dependencies, [SkiaSharp.NativeAssets.Linux](https://www.nuget.org/packages/SkiaSharp.NativeAssets.Linux/) should be installed. Also, installing of additional components is required as shown by the following commands:

```
apt-get update && apt-get install -y libfontconfig1
```

```
apt-get update && apt-get install -y --allow-unauthenticated libc6-dev libgdiplus libx11-dev
```

## Use GroupDocs.Assembly for .NET Standard 2.0 in macOS

In addition to GroupDocs.Assembly NuGet package dependencies, [runtime.osx.10.10-x64.CoreCompat.System.Drawing](https://www.nuget.org/packages/runtime.osx.10.10-x64.CoreCompat.System.Drawing/) should be installed.

## Use GroupDocs.Assembly for .NET Standard 2.0 in Docker

Additional steps required to use GroupDocs.Assembly for .NET Standard 2.0 in Docker correspond to a type of containers being used. For example, for Linux containers, you should take the same additional steps as per "Using GroupDocs.Assembly for .NET Standard 2.0 in Linux" section.
