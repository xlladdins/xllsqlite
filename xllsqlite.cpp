// xllsqlite.cpp - sqlite wrapper
#include <locale>
#include "xllsqlite.h"

using namespace xll;
/*
AddIn xai_sqlite(
    Documentation("Sqlite3 wrapper")
);
*/

#define SQLITE_OPEN_URL "https://sqlite.org/c3ref/open.html"

XLL_CONST(LONG, SQLITE_OPEN_READONLY, SQLITE_OPEN_READONLY, "Read only access.", CATEGORY, SQLITE_OPEN_URL);
XLL_CONST(LONG, SQLITE_OPEN_READWRITE, SQLITE_OPEN_READWRITE, "Read and write access.", CATEGORY, "Read and write access.");
XLL_CONST(LONG, SQLITE_OPEN_CREATE, SQLITE_OPEN_CREATE, "Create database if it does not already exist.", CATEGORY, SQLITE_OPEN_URL);
XLL_CONST(LONG, SQLITE_OPEN_URI, SQLITE_OPEN_URI, "Open using Universal Resource Identifier.", CATEGORY, SQLITE_OPEN_URL);
XLL_CONST(LONG, SQLITE_OPEN_MEMORY, SQLITE_OPEN_MEMORY, "Open data base in memory.", CATEGORY, SQLITE_OPEN_URL);
XLL_CONST(LONG, SQLITE_OPEN_NOMUTEX, SQLITE_OPEN_NOMUTEX, "Do not use mutal exclusing when accessing database.", CATEGORY, SQLITE_OPEN_URL);
XLL_CONST(LONG, SQLITE_OPEN_FULLMUTEX, SQLITE_OPEN_FULLMUTEX, "Use mutal exclusing when accessing database.", CATEGORY, SQLITE_OPEN_URL);

AddIn xai_sqlite_db(
    Function(XLL_HANDLE, "xll_sqlite_db", "SQLITE.DB")
    .Arguments({
        Arg(XLL_CSTRING4, "file", "is the name of the sqlite3 database to open."),
        Arg(XLL_SHORT, "flags", "is an optional set of flags from the SQLITE_OPEN_* enumeration to use when opening the database. Default is SQLITE_OPEN_READONLY."),
        })
    .Uncalced()
    .FunctionHelp("Return a handle to a sqlite3 database.")
    .Category("SQLITE")
    .Documentation("Call sqlite3_open_v2 with flags SQLITE_OPEN_READONLY.")
);
HANDLEX WINAPI xll_sqlite_db(const char* file, SHORT flags)
{
#pragma XLLEXPORT
    HANDLEX h = INVALID_HANDLEX;

    try {
        if (flags == 0)
            flags = SQLITE_OPEN_READONLY;

        handle<sqlite::db> h_(new sqlite::db(file, flags));
        h = h_.get();
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return h;
}

AddIn xai_sqlite_exec(
    Function(XLL_LPOPER, "?xll_sqlite_exec", "SQLITE.EXEC")
    .Arguments({
        Arg(XLL_HANDLE, "handle", "is the sqlite3 database handle returned by SQLITE.DB."),
        Arg(XLL_CSTRING4, "sql", "is the SQL query to execute on the database."),
        Arg(XLL_BOOL, "?headers", "is an optional argument to specify if headers should be included. Default is false."),
        })
    .FunctionHelp("Return the result of executing a SQL command on a database.")
    .Category("SQLITE")
    .Documentation("")
);
LPOPER WINAPI xll_sqlite_exec(HANDLEX h, const char* sql, BOOL headers)
{
#pragma XLLEXPORT
    static OPER o;

    try {
        handle<sqlite::db> h_(h);
        ensure (h_.ptr());
        o = sqlite_range(*h_, sql, headers);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return &o;
}

#if 0
AddIn xai_sqlite_table_info(
    Function(XLL_LPOPER, "?xll_sqlite_table_info", "SQLITE.TABLE.INFO")
    .Arg(XLL_HANDLE, "handle", "is the sqlite3 database handle returned by SQLITE.DB.")
    .Arg(XLL_CSTRING4, "table", "is the name of a table in the database.")
    .Arg(XLL_BOOL, "?headers", "is an optional argument to specify if headers should be included. Default is false.")
    .FunctionHelp("Return the cid, name, type, default value, and whether it is a primary key.")
    .Category("SQLITE")
    .Documentation("")
);
LPOPER WINAPI xll_sqlite_table_info(handlex h, const char* table, BOOL headers)
{
#pragma XLLEXPORT
    static OPER o;

    try {
        std::string sql = "PRAGMA table_info(";
        sql.append(table);
        sql.append(")");
        handle<sqlite::db> h_(h);
        ensure(h_.ptr());
        o = sqlite_range(*h_, sql.c_str(), headers);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return &o;
}

#ifdef _DEBUG

xll::test test_sqlite_range([]{
    char dir[1024];
    GetCurrentDirectoryA(1023, dir);
    const char* file = "C:/Users/kal/Source/Repos/keithalewis/xllsqlite/chinook.db";
    const char* sql = "PRAGMA table_info(artists);";
    sqlite::db db(file);
    sqlite::db::stmt stmt(db);
    OPER o;
    o = sqlite_range(db, sql, true);
    o = o;

});

#endif // _DEBUG
#endif // 0
