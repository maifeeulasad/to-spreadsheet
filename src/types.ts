/**
 * @fileoverview Shared runtime enums for the to-spreadsheet library.
 *
 * These enums are used as runtime values by both the writer (generate-excel,
 * util, worksheet/style generators) and the public entry point (index). Keeping
 * them in a dependency-free leaf module breaks what would otherwise be a circular
 * import between `index` and the writer modules, which made import-time evaluation
 * order fragile (notably under ESM/test transforms). `index` re-exports every one
 * of these, so the public API is unchanged.
 *
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

/**
 * Enum representing different cell types in Excel
 * @enum {string}
 */
enum ICellType {
  /** String cell type - contains text values */
  string = "s",
  /** Number cell type - contains numeric values */
  number = "n",
  /** Date cell type - contains date values */
  date = "d",
  /** Skip cell type - represents empty/skipped cells */
  skip = "skip",
  /** Equation cell type - contains Excel formulas */
  equation = "equation",
}

/**
 * Enum representing different border styles available in Excel
 * @enum {string}
 */
enum BorderStyle {
  /** No border */
  none = "none",
  /** Thin border line (default) */
  thin = "thin",
  /** Medium thickness border line */
  medium = "medium",
  /** Thick border line */
  thick = "thick",
  /** Double border line */
  double = "double",
  /** Dotted border line */
  dotted = "dotted",
  /** Dashed border line */
  dashed = "dashed",
}

/**
 * Enum representing horizontal alignment options for Excel cells
 * @enum {string}
 */
enum HorizontalAlignment {
  /** General alignment (Excel default) */
  general = "general",
  /** Left alignment */
  left = "left",
  /** Center alignment */
  center = "center",
  /** Right alignment */
  right = "right",
  /** Fill alignment */
  fill = "fill",
  /** Justify alignment */
  justify = "justify",
  /** Center across selection */
  centerContinuous = "centerContinuous",
  /** Distributed alignment */
  distributed = "distributed",
}

/**
 * Enum representing vertical alignment options for Excel cells
 * @enum {string}
 */
enum VerticalAlignment {
  /** Top alignment */
  top = "top",
  /** Center alignment */
  center = "center",
  /** Bottom alignment */
  bottom = "bottom",
  /** Justify alignment */
  justify = "justify",
  /** Distributed alignment */
  distributed = "distributed",
}

export { ICellType, BorderStyle, HorizontalAlignment, VerticalAlignment };
