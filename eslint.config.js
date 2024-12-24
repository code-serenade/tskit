import globals from "globals";
import pluginJs from "@eslint/js";
import tseslint from "typescript-eslint";

/** @type {import('eslint').Linter.Config[]} */
export default [
  { files: ["**/*.{ts}"] },
  { languageOptions: { parser: tsParser, globals: globals.browser } },
  pluginJs.configs.recommended,
  ...tseslint.configs.recommended,
];
