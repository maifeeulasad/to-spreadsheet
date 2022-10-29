# Hi
I appreciate, that you want to help us grow. I think we can help you get started.

# No TS?
If you are only familiar with JS, I, Maifee Ul Asad, personally will help you learn TS. TS is awesome. We don't have to maintain another repo for typing or that `.d.ts` file separately.
We are trying to make the quality of the code higher.

# Files
## Current organization
There is a few simple set of rules:
 - If the file name is `index.ts` or `generate*ts`, maybe it contains some function that will be exported. And is responsible for generating at least more than one part of the spreadsheet.
 - Else, it is already assigned to generate a specific part(to be a more specific file) of that spreadsheet.
## Why we did do it?
So how/why these parts were defined? We simply took a simple excel(.xlsx) file and extracted it. And kept that file structure. Read more here: https://github.com/maifeeulasad/to-spreadsheet/discussions/1
## How are we going to maintain it?
 - 

# Why it's a long way to go?
 - Let's take a look at something, which is working (almost): https://github.com/maifeeulasad/to-spreadsheet/blob/7295c884dfbbc20ac9ec0c456a14535adb63928c/src/xl/worksheets/sheet1.xml.ts#L37-L50. Now this works, this generates sheet seamlessly. But the issue is it doesn't support any style, color, alignment, etc. There are many arguments for that, and we have to implement those.
 - The other thing, they are completely missing. Say we have made, this line static: https://github.com/maifeeulasad/to-spreadsheet/blob/7295c884dfbbc20ac9ec0c456a14535adb63928c/src/generate-excel.ts#L21. But there can and will be more than one sheet, so we have to take care of these too.


**We are way too noob, in this, but we can achieve something great for sure.**
Thanks.
