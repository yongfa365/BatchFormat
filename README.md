[![AppVeyor CI](https://ci.appveyor.com/api/projects/status/github/yongfa365/BatchFormat?svg=true)](https://ci.appveyor.com/project/yongfa365/BatchFormat/branch/master)

## Introduction

`BatchFormat` is a Visual Studio extensions.

It can batch format `*.cs` file on these **areas**: Solution, Project, Folder, File.

## Function:

Remove, Sort, Using, Format Document, and more combinations.

## Support:

Visual Studio 2010, 2012, 2013, 2015.

## Building from Source:

Now this project can be edited in VS 2015 Community, Professional or Enterprise.
To build with older versions, you will need to install [Roslyn](https://github.com/dotnet/roslyn) package:

```powershell
Install-Package Microsoft.Net.Compilers

# Note that in this case, the editor will show the C#6
# syntax as invalid, but the compiler will continue to work.
```

## Screenshot:

![Preview 1](http://i1.visualstudiogallery.msdn.s-msft.com/a7f75c34-82b4-4357-9c66-c18e32b9393e/image/file/52181/1/preview.png)

![Preview 2](https://i1.visualstudiogallery.msdn.s-msft.com/a7f75c34-82b4-4357-9c66-c18e32b9393e/image/file/179382/1/batchformat.option.png)

## src debug：
 http://stackoverflow.com/questions/24653486/how-to-debug-a-vsix-project

- Open a Visual Studio instance for development. Load up the solution.
- Make some changes, build.
- Install the .vsix (double clicking the vsix file in \bin\debug)
- Launch a new Instance of Visual Studio for testing.
- In the Dev Visual Studio, Debug -> Attach to Process and select devenv.exe
- In the Dev Visual Studio, set some break points.
- In the Test Visual Studio, create / open new solution and start test coding.
