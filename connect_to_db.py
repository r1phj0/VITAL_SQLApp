#!/usr/bin/env python
# coding: utf-8

def get_connection():
    # In[ ]:
    import cx_Oracle

    # In[ ]:
    print("Connecting to oracle...")
    dsn = """(DESCRIPTION =
        (LOAD_BALANCE=on)
        (FAILOVER=on)
        (ADDRESS=(PROTOCOL=tcp)(HOST=eagnmnmep0b3)(PORT=1521))
        (ADDRESS=(PROTOCOL=tcp)(HOST=eagnmnmep0b4)(PORT=1521))
        (ADDRESS=(PROTOCOL=tcp)(HOST=eagnmnmep0b5)(PORT=1521))
        (CONNECT_DATA=(SERVICE_NAME=pvital.usps.gov)))"""
    conn = cx_Oracle.connect(
        user='DB_R1PHJ0',
        password='Postalservice08',
        dsn=dsn,
        encoding="UTF-8"
    )
    print("Connection successful!")
    return conn


if __name__ == "__main__":
    conn = get_connection()
