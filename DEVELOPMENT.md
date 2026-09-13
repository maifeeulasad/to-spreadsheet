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

# Building and testing

We use [pnpm](https://pnpm.io/). Install once with `pnpm i`.

```
pnpm build        # compile TypeScript to lib/ (tsc)
pnpm test         # run the vitest suite once
pnpm test:watch   # re-run tests on change
```

## How the code is split

- **Writer** — `generate-excel.ts` (orchestration), `util.ts` (cell helpers), and
  the `xl/`, `docProps/`, `_rels/` generators that each emit one part of the
  `.xlsx` package. These turn an `IPage[]` into the zipped OOXML file.
- **Reader** — `read-excel.ts` (`readExcel`, unzips OOXML via JSZip and parses the
  worksheet/sharedStrings/styles parts back into values) and `read-csv.ts`
  (`parseCsv`, a dependency-free RFC 4180 parser). Both run in Node and the browser.
- **Shared** — `types.ts` holds the runtime enums (`ICellType`, `BorderStyle`,
  alignment) imported by both sides. Keep enums here, not in `index.ts`, so the
  writer modules never import `index` at runtime (that cycle broke test-time
  evaluation order).

## Writing tests

Tests live next to the code as `*.test.ts` (excluded from the published build).
The reader tests build a workbook in memory with the writer and read it straight
back — the round-trip is the strongest guarantee that the two halves agree, so
prefer adding to it when you touch either side.
