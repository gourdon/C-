README for C# In Depth source code
----------------------------------

The source code for C# In Depth is laid out in the following manner:

MiscUtil.dll:  Required for most solutions, to allow easy starting
               of individual programs within projects.

Chapter12:     Source code for chapter 12 (various LINQ providers)
Snippy:        Source code for the "Snippy" snippet compiler
Snippets:      The snippets themselves
Databases:     Database files
OtherChapters: Code for chapters 1-11, 13-15

Chapter 12 uses multiple projects to lay out their code;
the other chapters all have a single project per chapter (or one for
VS2008 and one for VS2010 only), and so can easily live within a
single solution.

There are two solution files in the "OtherChapters" and "Chapter12"
directories, one for Visual Studio 2008 and one for Visual Studio
2010. The Visual Studio 2010 solution contain all projects, whereas
the Visual Studio 2008 solutions don't have any projects which
require C# 4 or .NET 4.0.

In some cases the code is slightly different from that in the book, in order
to use namespaces and type names which are more appropriate for types within
a full solution.

Each project uses MiscUtil as a launcher - set a project as the startup
project, run it, and you will be presented with options of which program to
run from within that project.
