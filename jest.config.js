/** @type {import('ts-jest/dist/types').InitialOptionsTsJest} */
module.exports = {
    globals: {
        "ts-jest": {
            tsconfig: "<rootDir>/tests/tsconfig.json",
        },
    },
    collectCoverageFrom: ["<rootDir>/src/**/*.ts"],
    testMatch: [
        "<rootDir>/tests/**/?(*.)(spec|test).(ts|tsx)",
    ],
    preset: "ts-jest",
};
