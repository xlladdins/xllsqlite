// xllsqlite.cpp - sqlite wrapper
#include <locale>
#include "xllsqlite.h"

using namespace xll;
using xcstr = traits<XLOPERX>::xcstr;

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

AddIn xai_sqlite_open(
    Function(XLL_HANDLE, "xll_sqlite_open", "\\SQLITE.OPEN")
    .Arguments({
        Arg(XLL_CSTRING4, "file", "is the name of the sqlite3 database to open."),
        Arg(XLL_LONG, "flags", "is an optional set of flags from the SQLITE_OPEN_* enumeration to use when opening the database. Default is SQLITE_OPEN_READONLY."),
        })
    .Uncalced()
    .FunctionHelp("Return a handle to a sqlite3 database.")
    .Category(CATEGORY)
    .HelpTopic("https://www.sqlite.org/c3ref/open.html")
    //.Documentation("")
);
HANDLEX WINAPI xll_sqlite_open(const char* file, LONG flags)
{
#pragma XLLEXPORT
    HANDLEX h = INVALID_HANDLEX;

    try {
        if (flags == 0)
            flags = SQLITE_OPEN_READONLY;

        handle<sqlite::open> h_(new sqlite::open(file, flags));
        h = h_.get();
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return h;
}

/*
CREATE.TABLE(table-name, {name, constraint(type-name);...})
CREATE.TEMP.TABLE
CREATE.TABLE.IF_NOT_EXISTS
CREATE.TEMP.TABLE.IF_NOT_EXISTS

CONSTRAINT(name, ...)
PRIMARY_KEY[.ASC|.DESC](clause)[.AUTOINCREMENT]
    ON_CONFLICT_ROLLBACK
    ON_CONFLICT_ABORT
    ...

HAVING(expr, GROUP_BY({expr,...}, WHERE(expr, FROM(table, SELECT({column,...})))))
SELECT.ALL
SELECT.DISTINCT
*/

/*
AddIn xai_range(
    Function(XLL_LPOPER4, "xll_range", "RANGE")
    .Arguments({
        Arg(XLL_HANDLE, "handle", "is a handle to a range.")
        })
    .Category("XLL")
    .FunctionHelp("Return the range corresponding to a handle.")
);
*/
//struct SELECT : public OPER4 {};

AddIn xai_sql_select(
    Function(XLL_LPOPER4, "xll_sql_select", "SQL.SELECT") // .ALL, .DISTINCT
    .Arguments({
        Arg(XLL_LPOPER4, "columns", "is a range of the columns to return."),
        })
    .Category(CATEGORY)
    .FunctionHelp("Return SQL SELECT statement.")
    .HelpTopic("https://www.sqlite.org/syntax/select-core.html")
);
LPOPER4 WINAPI xll_sql_select(const LPOPER4 pcols)
{
#pragma XLLEXPORT
    static OPER4 result;
    result = ErrValue4;

    try {
        OPER4 sel("SELECT ");
        OPER4 comma("");
        for (const auto& col : *pcols) {
            ensure(col.is_str());
            sel &= comma;
            sel &= col;
            comma = ", ";
        }
        result.swap(sel);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return &result;
}

AddIn xai_sql_from(
    Function(XLL_LPOPER4, "xll_sql_from", "SQL.FROM")
    .Arguments({
        Arg(XLL_CSTRING4, "table", "is the table to select from."),
        Arg(XLL_LPOPER4, "select", "is a SELECT statement."),
        })
        .Category(CATEGORY)
    .FunctionHelp("Return SQL from statement.")
    .HelpTopic("https://www.sqlite.org/syntax/select-core.html")
);
LPOPER4 WINAPI xll_sql_from(const char* table, const LPOPER4 psel)
{
#pragma XLLEXPORT
    static OPER4 result;
    result = ErrNA4;

    try {
        result = *psel;
        result.resize(result.size(), 1);
        result.push_bottom(OPER4("FROM ").append(table));
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return &result;
}

AddIn xai_sql_where(
    Function(XLL_LPOPER4, "xll_sql_where", "SQL.WHERE")
    .Arguments({
        Arg(XLL_CSTRING4, "expr", "is an expresion."),
        Arg(XLL_LPOPER4, "from", "is a FROM statement."),
        })
        .Category(CATEGORY)
    .FunctionHelp("Return SQL where statement.")
    .HelpTopic("https://www.sqlite.org/syntax/select-core.html")
);
LPOPER4 WINAPI xll_sql_where(const char* expr, const LPOPER4 psel)
{
#pragma XLLEXPORT
    static OPER4 result;
    result = ErrNA4;

    try {
        result = *psel;
        result.resize(result.size(), 1);
        result.push_bottom(OPER4("WHERE ").append(expr));
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return &result;
}

AddIn xai_sql_group_by(
    Function(XLL_LPOPER4, "xll_sql_group_by", "SQL.GROUP_BY")
    .Arguments({
        Arg(XLL_LPOPER4, "exprs", "is a range of expresions."),
        Arg(XLL_LPOPER4, "where", "is a WHERE statement."),
        })
        .Category(CATEGORY)
    .FunctionHelp("Return SQL GROUP BY statement.")
    .HelpTopic("https://www.sqlite.org/syntax/select-core.html")
);
LPOPER4 WINAPI xll_sql_group_by(const LPOPER4 pexprs, const LPOPER4 psel)
{
#pragma XLLEXPORT
    static OPER4 result;
    result = ErrNA4;

    try {
        OPER4 gb("GROUP BY ");
        OPER4 comma("");
        for (const auto& expr : *pexprs) {
            ensure(expr.is_str());
            gb &= comma;
            gb &= expr;
            comma = ", ";
        }
        result = *psel;
        result.resize(result.size(), 1);
        result.push_bottom(gb);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return &result;
}

// HAVING(expr)
// ORDER_BY(exprs)
// LIMIT(expr,...

AddIn xai_create_table(
    Function(XLL_HANDLE, "xll_create_table", "SQLITE.CREATE_TABLE")
    .Arguments({
        Arg(XLL_HANDLE, "handle", "is a handle to a database."),
        Arg(XLL_CSTRING4, "table", "is the name of the table."),
        Arg(XLL_LPOPER4, "names", "is an array of column names."),
        Arg(XLL_LPOPER4, "types", "is an array of SQL types."),
        Arg(XLL_LPOPER4, "_contraints", "is an optional array of contraints."),
        })
    .Category(CATEGORY)
    .FunctionHelp("Create a table in a database and return handle.")
);
HANDLEX WINAPI xll_create_table(HANDLEX h, const char* table, 
    const LPOPER4 pnames, const LPOPER4 ptypes, const LPOPER4 pconstraints)
{
#pragma XLLEXPORT
    try {
        ensure(pnames->size() == ptypes->size());
        ensure(pconstraints->is_missing() or pnames->size() == pconstraints->size());

        handle<sqlite::open> h_(h);
        ensure(h_.ptr());

        std::string ct("CREATE TABLE ");
        ct.append(table);
        ct.append(" (");
  
        std::string comma = "";
        const OPER4& name(*pnames);
        const OPER4& type(*ptypes);
		for (unsigned i = 0; i < pnames->size(); ++i) {
			ensure(name[i].is_str());
			ensure(type[i].is_str());
			ct.append(comma);
            ct.append(name[i].val.str + 1, name[i].val.str[0]);
            ct.append(" ");
            ct.append(type[i].val.str + 1, type[i].val.str[0]);
            if (!pconstraints->is_missing()) {
                const auto& ci = index(*pconstraints, i);
                ensure(ci.is_str());
                if (ci) {
                    ct.append(" ");
                    ct.append(ci.val.str + 1, ci.val.str[0]);
                }
            }

            comma = ", ";
        }
        ct.append(")");

        sqlite_exec(*h_, ct.c_str());
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return h;
}

AddIn xai_sqlite_exec(
    Function(XLL_LPOPER4, "xll_sqlite_exec", "SQLITE.EXEC")
    .Arguments({
        Arg(XLL_HANDLE, "handle", "is the sqlite3 database handle returned by SQLITE.OPEN."),
        Arg(XLL_LPOPER4, "sql", "is the SQL query to execute on the database."),
        Arg(XLL_BOOL, "_headers", "is an optional argument to specify if headers should be included. Default is false."),
        })
    .FunctionHelp("Return the result of executing a SQL command on a database.")
    .Category(CATEGORY)
    .HelpTopic("https://www.sqlite.org/c3ref/exec.html")
    .Documentation("")
);
LPOPER4 WINAPI xll_sqlite_exec(HANDLEX h, const LPOPER4 psql, BOOL headers)
{
#pragma XLLEXPORT
    static OPER4 o;
    o = ErrNA4;

    try {
        std::string sql;
        for (const auto& s : *psql) {
            ensure(s.is_str());
            sql.append(s.val.str + 1, s.val.str[0]);
            sql.append(" ");
        }

        handle<sqlite::open> h_(h);
        ensure (h_.ptr());
        
        o = sqlite_exec(*h_, sql.c_str(), headers);
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
