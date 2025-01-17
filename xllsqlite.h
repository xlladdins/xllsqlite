// xllsqlite.h - sqlite3 wrapper
#pragma once
#include "sqlite3.h"
#include "xll/xll/xll.h"

#define CATEGORY "SQLite"

namespace sqlite {

    enum class Type {
        Integer = SQLITE_INTEGER,
        Float = SQLITE_FLOAT,
        Text = SQLITE_TEXT,
        Blob = SQLITE_BLOB,
        Null = SQLITE_NULL,
    };

    class value {
        sqlite3_value* val;
    public:
        value()
            : val(sqlite3_value_dup(nullptr))
        { }
        value(const value& v)
            : val(sqlite3_value_dup(v.val))
        { }
        value& operator=(const value& v)
        {
            if (this != &v) {
                sqlite3_value_free(val);
                val = sqlite3_value_dup(v.val);
            }

            return *this;
        }
        value(value&& v) noexcept
            : val(v.val)
        {
            v.val = sqlite3_value_dup(0);
        }
        value& operator=(value&& v) noexcept
        {
            std::swap(val, v.val);

            return *this;
        }
        ~value()
        {
            sqlite3_value_free(val);
        }

        int type() const
        {
            return sqlite3_value_type(val);
        }

        int bytes() const
        {
            return sqlite3_value_bytes(val);
        }
    };

    // Sqlite converts wide strings to UTF-8 so we avoid *16* functions.
    class open {
        sqlite3* pdb;
    public:
        open(const char* file, int flags = SQLITE_OPEN_READONLY)
        {
            if (SQLITE_OK != sqlite3_open_v2(file, &pdb, flags, 0))
                throw std::runtime_error(sqlite3_errmsg(pdb));
        }
        open(const open&) = delete;
        open& operator=(const open&) = delete;
        ~open()
        {
            sqlite3_close(pdb);
        }
        // for use in sqlite3_* functions
        operator sqlite3*() {
            return pdb;
        }
        class stmt {
            sqlite::open& db;
            sqlite3_stmt* pstmt;
            const char* tail_;
        public:
            stmt(sqlite::open& db)
                : db(db), pstmt(nullptr), tail_(nullptr)
            { }
            stmt(const stmt&) = delete;
            stmt& operator=(const stmt&) = delete;
            ~stmt()
            {
                sqlite3_finalize(pstmt);
            }
            // for use in sqlite3_* functions
            operator sqlite3_stmt*()
            {
                return pstmt;
            }
            const char* errmsg() const
            {
                return sqlite3_errmsg(db);
            }
            int prepare(const char* sql, int nsql = -1)
            {
                return sqlite3_prepare_v2(db, sql, nsql, &pstmt, &tail_);
            }
            const char* tail() const
            {
                return tail_;
            }
            int bind(int col, int i)
            {
                return sqlite3_bind_int(pstmt, col, i);
            }
            int bind(int col, sqlite_int64 i)
            {
                return sqlite3_bind_int64(pstmt, col, i);
            }
            int bind(int col, double d)
            {
                return sqlite3_bind_double(pstmt, col, d);
            }
            // Do not make a copy of text by default.
            int bind(int col, const char* t, int n = -1, void(*dealloc)(void*) = SQLITE_STATIC)
            {
                return sqlite3_bind_text(pstmt, col, t, n, dealloc);
            }
        };
    };
}

// convert wide string to UTF-8
inline std::string narrow(const wchar_t* ws, int ns = -1)
{
    std::string s;

    int n;
    if (ns == -1) {
        n = WideCharToMultiByte(CP_UTF8, 0, ws, ns, 0, 0, 0, 0);
    }
    else {
        n = ns;
    }
    s.reserve(n);
    
    int rc = WideCharToMultiByte(CP_UTF8, 0, ws, ns, &s[0], n, 0, 0);
    ensure (rc != 0);

    return s;
}

// Sqlite type of oper.
inline const char* sqlite_type(const xll::OPER4& o)
{
    switch (o.type()) {
    case xltypeNum:
        return "REAL";
    case xltypeBigData:
        return "BLOB";
    }

    return "TEXT";
}

// Works like sqlite3_exec but returns an OPER.
inline xll::OPER4 sqlite_exec(sqlite::open& db, const char* sql, bool header = false)
{
    xll::OPER4 o;

    sqlite::open::stmt stmt(db);
    int rc = stmt.prepare(sql);
    if (SQLITE_OK != rc)
        throw std::runtime_error(stmt.errmsg());
    
    rc = sqlite3_step(stmt);
    if (rc != SQLITE_ROW && rc != SQLITE_DONE)
        throw std::runtime_error(stmt.errmsg());
   
    int n = sqlite3_column_count(stmt);
    if (header) {
        xll::OPER4 head(1, n);
        for (int i = 0; i < n; ++i) {
            head[i] = sqlite3_column_name(stmt, i);
        }
        o.push_back(head);
    }
    while (SQLITE_ROW == rc) {
        xll::OPER4 row(1, n);
        for (int i = 0; i < n; ++i) {
            switch (sqlite3_column_type(stmt, i)) {
            case SQLITE_FLOAT:
                row[i] = sqlite3_column_double(stmt, i);
                break;
            case SQLITE_TEXT:
                row[i] = (char*)sqlite3_column_text(stmt, i);
                break;
            case SQLITE_INTEGER:
                row[i] = sqlite3_column_int(stmt, i);
                break;
            case SQLITE_NULL:
                row[i] = xll::ErrNull4;
                break;
            default:
                row[i] = xll::ErrNA4;
            }
        }
        o.push_back(row);
        rc = sqlite3_step(stmt);
    };

    return o;
}

