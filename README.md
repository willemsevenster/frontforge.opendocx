# Frontforge OpenDocx
A simple wrapper for DocumentFormat.OpenXml word processing documents (MS Word and similar).

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
public class MyDocument : WordDocument 
{
    public void Create(string fileName) {
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
            .BottomBorder();

        tbl.Add(
            Row(
                Cell(Par("Cell 0, 0").Bold()), 
                Cell(Par("Cell 0, 1"))
            ),
            Row(
                Cell(Par("Cell 1, 0").Bold()),
                Cell(Par("Cell 1, 1"))
            )
        );

        // add the table to the section
        section.Add(tbl);

        // add the section to the doument
        AddSection(section);

        // save to file
        using (var fileStream = new FileStream(fileName, FileMode.Create))
        {
            Save(fileStream);
            fileStream.Flush();
        }

        // you can also return the WordDocument object that can be saved or streamed
        // e.g.: return this;
    }
}
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

