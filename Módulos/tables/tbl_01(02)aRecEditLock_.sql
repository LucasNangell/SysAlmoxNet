-- Tabela: tbl_01(02)aRecEditLock_
-- Registros: 2

CREATE TABLE tbl_01(02)aRecEditLock_ (
    RecLockID INTEGER AUTOINCREMENT,
    LockedRecID INTEGER,
    RecSource VARCHAR(255),
    UserLoginID INTEGER,
    EditStartTime DATETIME DEFAULT Now()
);

ALTER TABLE tbl_01(02)aRecEditLock_ ADD PRIMARY KEY (RecLockID);
CREATE INDEX RecID ON tbl_01(02)aRecEditLock_ (LockedRecID);
CREATE INDEX RecLockID ON tbl_01(02)aRecEditLock_ (RecLockID);

