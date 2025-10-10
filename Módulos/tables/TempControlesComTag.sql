-- Tabela: TempControlesComTag
-- Registros: 1399

CREATE TABLE TempControlesComTag (
    ID INTEGER AUTOINCREMENT,
    NomeFormulario VARCHAR(255),
    NomeControle VARCHAR(255),
    TipoControle VARCHAR(255),
    FonteControle VARCHAR(255),
    Rotulo VARCHAR(255),
    Tag VARCHAR(255),
    Visivel BIT,
    Habilitado BIT,
    Topo INTEGER,
    Esquerda INTEGER,
    Largura INTEGER,
    Altura INTEGER,
    DataHora DATETIME
);

ALTER TABLE TempControlesComTag ADD PRIMARY KEY (ID);

