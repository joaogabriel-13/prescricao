export interface Planilha {
    nome: string;
    dados: Array<Registro>;
}

export interface Registro {
    [key: string]: any;
}