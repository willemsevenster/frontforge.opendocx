# Frontforge OpenDocx
A simple wrapper for DocumentFormat.OpenXml word processing documents (MS Word and similar).

## Status
[![Quality Gate Status](https://sonarqube.frontforge.com/api/project_badges/measure?project=willemsevenster_frontforge.opendocx&metric=alert_status)](https://sonarqube.frontforge.com/dashboard?id=willemsevenster_frontforge.opendocx)

## Overview
Allows easier creation of OpenXML wordprocessing documents using a fluent API. This is a .Net Standard 2.0 library and depends on the DocumentFormat.OpenXml library. This library was inspired by the [ReportDotNet](https://github.com/mifopen/ReportDotNet) project on Github.

## Dependencies
1. .NETStandard 2.0
2. [DocumentFormat.OpenXml (>= 2.9.1)](https://www.nuget.org/packages/DocumentFormat.OpenXml/)

## Usage
The library exposes an abstract class `WordDocument` that implements fluent methods to assist in creating sections, parargraphs, checkboxes and tables (with rows and cells).

There is a lot of functionality that must still be added, but this library is already useful for creating fairly complex document layouts.

### Getting started
Available on [nuget.org](https://www.nuget.org/packages/Frontforge.OpenDocx.Core/)

#### Install using Nuget
`Install-Package Frontforge.OpenDocx.Core`

#### Install using .Net CLI
`dotnet add package Frontforge.OpenDocx.Core`

#### Creating a basic wordprocessing document
See the `/examples` folder for more complete examples.

```c#
    internal class BasicSample
        : WordDocument
    {
        #region implementation

        public static BasicSample Create()
        {
            return new BasicSample().BuildDoc();
        }

        private BasicSample BuildDoc()
        {
            // create a new section and set section properties
            var section = Section()
                .PageSize(PageSize.A4)
                .PageMargins(PredefinedPageMargins.Narrow);

            // add a paragraph
            var par = Par("This is a simple paragraph that is bold and center aligned with a " +
                          "16pt font size and a 6pt spacing before the paragraph.",
                    HorizontalAlignment.Center)
                .SpacingBefore(6)
                .Bold()
                .FontSize(16);

            section.Add(par);

            // add a table
            var tbl = Table()
                .Width(new Unit(100, UnitType.pct)) // 100% width
                .TopBorder()
                .BottomBorder();

            tbl.Add(
                Row(
                    Cell(Par("Cell 0, 0").Bold()).Width(30, UnitType.pct),
                    Cell(Par("Cell 0, 1"))
                ),
                Row(
                    Cell(Par("Cell 1, 0").Bold()),
                    Cell(Par("Cell 1, 1"))
                )
            );

            // add the table to the section
            section.Add(tbl);

            // add the section to the document
            AddSection(section);

            return this;
        }

        #endregion
    }
```
This document can be created and saved to disk, or any other `Stream` as follows:

```c#
        // save to file
        using (var fileStream = new FileStream(FileNameBasicDoc1, FileMode.Create))
        {
            BasicSample.Create().Save(fileStream);
            fileStream.Flush();
        }

        // open the file
        var process = new Process
        {
            StartInfo =
            {
                UseShellExecute = true,
                FileName = FileNameBasicDoc1
            }
        };

        process.Start();
```

## Contributing
Any and all contributions are welcomed. Please follow this [excellent guide](https://akrabat.com/the-beginners-guide-to-contributing-to-a-github-project/#summary). __Note:__ This project uses [git-flow](http://nvie.com/posts/a-successful-git-branching-model/). 

The basic steps for contributing are:
1. Fork the project & clone locally.
2. Create an upstream remote and sync your local copy before you branch.
3. Branch for each separate piece of work.
4. Do the work, write good commit messages, and read the [CONTRIBUTING](./CONTRIBUTING.md) file.
5. Push to your origin repository.
6. Create a new PR in GitHub.
7. Respond to any code review feedback.

