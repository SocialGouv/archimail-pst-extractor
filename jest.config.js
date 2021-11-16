// /** @type {import('ts-jest/dist/types').InitialOptionsTsJest} */
// const path = require("path");
// const fs = require("fs");
// const { pathsToModuleNameMapper } = require("ts-jest/utils");

// const tsconfig = JSON.parse(
//     fs.readFileSync(path.resolve(__dirname, "tests", "tsconfig.json"), {
//         encoding: "utf-8",
//     })
// );

// const moduleNameMapper = {
//     ...pathsToModuleNameMapper(tsconfig.compilerOptions.paths, {
//         prefix: "<rootDir>",
//     }),
// };

module.exports = {
    collectCoverageFrom: ["<rootDir>/src/**/*.ts"],
    globals: {
        "ts-jest": {
            tsconfig: "<rootDir>/tests/tsconfig.json",
        },
    },
    moduleDirectories: ["<rootDir>/node_modules", "<rootDir>"],
    moduleFileExtensions: ["ts", "tsx", "js", "json"],
    // moduleNameMapper,
    preset: "ts-jest",
    testMatch: ["<rootDir>/tests/**/?(*.)(spec|test).(ts|tsx)"],
};
