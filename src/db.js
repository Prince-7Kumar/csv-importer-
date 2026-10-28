
const { Pool } = require("pg");

const requiredVariables = [
    "POSTGRES_HOST",
    "POSTGRES_DB",
    "POSTGRES_USER",
    "POSTGRES_PASSWORD",
];

for (const variable of requiredVariables) {
    if (!process.env[variable]) {
        throw new Error(
            `Missing ${variable} in .env. Check your database configuration.`
        );
    }
}

const sslEnabled =String(process.env.POSTGRES_SSL || "false").toLowerCase() === "true";//it 

const pool = new Pool({
    host: process.env.POSTGRES_HOST,
    port: Number(process.env.POSTGRES_PORT || 5432),
    database: process.env.POSTGRES_DB,
    user: process.env.POSTGRES_USER,
    password: process.env.POSTGRES_PASSWORD,

    ...(sslEnabled
        ? {
            ssl: {
                rejectUnauthorized: false,
            },
        }
        : {}),

    connectionTimeoutMillis: 10000,
    idleTimeoutMillis: 30000,
    max: 5,
});

module.exports = pool;
