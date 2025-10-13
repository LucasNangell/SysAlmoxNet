-- Tabela: tbl_01(03)bSysEnblDisblParams
-- Registros: 70

CREATE TABLE tbl_01(03)bSysEnblDisblParams (
    EnbleDsbleSetID INTEGER AUTOINCREMENT,
    sSysForM VARCHAR(255),
    sTriggerCtrl VARCHAR(255),
    sSysFormMode VARCHAR(255),
    sSetFocusToCtrl VARCHAR(255),
    sTweakbleCtrl VARCHAR(255),
    bEnable BIT DEFAULT 0,
    bVisible BIT DEFAULT 0,
    bLockCombo BIT DEFAULT 0,
    sAltTipText VARCHAR(255)
);

ALTER TABLE tbl_01(03)bSysEnblDisblParams ADD PRIMARY KEY (EnbleDsbleSetID);

