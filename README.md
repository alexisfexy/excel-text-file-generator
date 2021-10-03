# Excel To Text File Generation

Transforms excel sheet with desired text file names & content into text files. With ability to include text transformation too.


## Overview of Requirements:
1. excel file with columns for template names & columns for template contents (@Eddie Ciafardini anything similar to the file you gave me would work)
2. quick mapping of column indexes & any alterations you want to the text (e.g remove prefix, capitalize, etc). See below for example.

```{<Column Description>: {"file_name": (<index>, <callable of text alterations>),
                           "content": (<index>,, <callable of text alterations>}}

EXAMPLE: 
{
    "split":
        {"file_name": (2, lambda string: string.removeprefix("    SplitKey.").removesuffix(": ")),
           "content": (4, None)}
}```
