import JSZip from "jszip";
import { calcularNumeracao, lerMapaNumeracao, type MarcaNumerada } from "./numeracao";

/* ------------------------------------------------------------------ */
/* Tipos                                                               */
/* ------------------------------------------------------------------ */

export type Bloco = {
  /** XML bruto do elemento de corpo (w:p, w:tbl, ...) — preservado 100% */
  xml: string;
  /** texto do bloco, já com as entidades XML decodificadas */
  texto: string;
  tipo: "p" | "tbl" | "outro";
  temImagem: boolean;
  /** numeração automática do Word (w:numPr), quando existir */
  num?: MarcaNumerada | null;
};

export type Alternativa = {
  letra: string;
  /** rótulo original, ex.: "a) " ou "(B) " */
  rotulo: string;
  /** blocos que pertencem a esta alternativa (o primeiro contém o rótulo) */
  indices: number[];
};

export type Afirmacao = {
  rotulo: string;
  indices: number[];
};

export type Questao = {
  /** identidade estável da questão (nunca é o índice do array) */
  id: string;
  /** id do texto de apoio ao qual pertence, ou null quando é independente */
  supportId: string | null;
  numeroOriginal: number;
  posicaoOriginal: number;
  rotuloOriginal: string;
  /** true quando o número vem da numeração automática do Word (não do texto) */
  numeracaoAuto: boolean;
  preview: string;
  /** blocos do enunciado / texto de apoio / itens / imagens / tabelas */
  corpo: number[];
  alternativas: Alternativa[];
  /** afirmações ( ) V/F, quando a questão não tem alternativas */
  afirmacoes: Afirmacao[];
  /** blocos finais da questão que não devem ser movidos internamente */
  rodape: number[];
};

/** Um texto de apoio (âncora fixa) e todas as questões que dependem dele. */
export type Contexto = {
  /** identificador do grupo (texto_A, texto_B, independentes_1, ...) */
  id: string;
  /** blocos do texto de apoio compartilhado (vazio = grupo independente) */
  apoio: number[];
  /** índices em `questoes` que pertencem a este contexto */
  questoes: number[];
};

export type Diagnostico = {
  /** total de blocos (parágrafos/tabelas) lidos do documento */
  blocos: number;
  /** marcadores de numeração candidatos encontrados antes do filtro */
  candidatos: number;
  /** candidatos vindos da numeração automática do Word */
  candidatosAuto: number;
  questoes: number;
  textosApoio: number;
  independentes: number;
  imagens: number;
  tabelas: number;
  grupos: { id: string; questoes: string[] }[];
};


export type AnaliseProva = {
  /** todos os blocos do corpo do documento, na ordem original */
  blocos: Bloco[];
  /** blocos antes da primeira questão (cabeçalho da prova) */
  cabecalho: number[];
  questoes: Questao[];
  /** agrupamento texto de apoio → questões, na ordem original do documento */
  contextos: Contexto[];
  /** <w:sectPr> final, se existir (mantém colunas, margens, orientação) */
  sectPr: string;
  /** XML do documento antes de <w:body> e depois de </w:body> */
  prefixo: string;
  sufixo: string;
  /** bytes do arquivo original, para reempacotar preservando mídia */
  bytes: ArrayBuffer;
  diagnostico: Diagnostico;
};

/* ------------------------------------------------------------------ */
/* Utilidades de texto                                                 */
/* ------------------------------------------------------------------ */

export function decodificar(s: string) {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, "&");
}

/** Texto visível do bloco: junta runs, tabulações e quebras, decodificado. */
function textoDoBloco(xml: string) {
  const partes: string[] = [];
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/?>|<w:br\b[^>]*\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) partes.push(m[1] !== undefined ? m[1] : " ");
  return decodificar(partes.join("")).replace(/\u00a0/g, " ");
}

/* ------------------------------------------------------------------ */
/* 1. DOCX Parser — estrutura real do documento                        */
/* ------------------------------------------------------------------ */

/** Nome da tag de um elemento XML ("<w:p ...>" -> "w:p"). */
function nomeTag(el: string) {
  return el.slice(1).split(/[\s/>]/)[0] ?? "";
}

/** Escapa caracteres especiais para uso em RegExp. */
function escaparRegex(s: string) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Divide os filhos diretos de um trecho XML preservando o XML de cada um.
 * Usa uma pilha de profundidade e casa o nome COMPLETO da tag, para que
 * `</w:tblPr>` nunca seja confundido com o fechamento de `<w:tbl>`.
 */
function dividirElementos(interno: string): string[] {
  const blocos: string[] = [];
  let i = 0;
  while (i < interno.length) {
    const abre = interno.indexOf("<", i);
    if (abre < 0) break;
    const fechaTag = interno.indexOf(">", abre);
    if (fechaTag < 0) break;
    const conteudoTag = interno.slice(abre + 1, fechaTag);
    const nome = conteudoTag.split(/[\s/>]/)[0]!;
    if (conteudoTag.endsWith("/") || nome.startsWith("!") || nome.startsWith("?")) {
      blocos.push(interno.slice(abre, fechaTag + 1));
      i = fechaTag + 1;
      continue;
    }

    // Percorre apenas tags cujo nome é EXATAMENTE `nome` (limite de nome real).
    const re = new RegExp("<(/?)" + escaparRegex(nome) + "(?=[\\s/>])([^>]*)>", "g");
    re.lastIndex = fechaTag + 1;
    let profundidade = 1;
    let cursor = interno.length;
    let m: RegExpExecArray | null;
    while ((m = re.exec(interno))) {
      const ehFechamento = m[1] === "/";
      const autoFechada = (m[2] ?? "").endsWith("/");
      if (ehFechamento) profundidade--;
      else if (!autoFechada) profundidade++;
      if (profundidade === 0) {
        cursor = m.index + m[0].length;
        break;
      }
    }
    blocos.push(interno.slice(abre, cursor));
    i = cursor;
  }
  return blocos;
}


/** Conteúdo interno de um elemento XML (vazio quando é auto-fechado). */
function conteudoInterno(el: string) {
  const nome = nomeTag(el);
  const abertura = el.indexOf(">");
  if (abertura < 0 || el.endsWith("/>")) return "";
  const fim = el.lastIndexOf("</" + nome + ">");
  if (fim < 0) return "";
  return el.slice(abertura + 1, fim);
}

/** Divide os filhos diretos de <w:body> preservando o XML de cada um. */
function dividirCorpo(xml: string) {
  const iniBody = xml.indexOf("<w:body");
  const fimAbertura = xml.indexOf(">", iniBody);
  const fimBody = xml.lastIndexOf("</w:body>");
  if (iniBody < 0 || fimBody < 0) throw new Error("DOCX sem corpo de documento.");

  const prefixo = xml.slice(0, fimAbertura + 1);
  const sufixo = xml.slice(fimBody);
  const interno = xml.slice(fimAbertura + 1, fimBody);

  return { prefixo, sufixo, blocos: dividirElementos(interno) };
}

/* ------------------------------------------------------------------ */
/* 1b. Achatamento estrutural (tabelas, células, caixas de texto)       */
/* ------------------------------------------------------------------ */

/** Números de questão detectados numa lista de elementos. */
function numerosDeQuestao(elementos: string[]): number[] {
  const nums: number[] = [];
  for (const el of elementos) {
    if (nomeTag(el) !== "w:p") continue;
    const q = detectarQuestao(textoDoBloco(el));
    if (q && !q.fraco) nums.push(q.numero);
  }
  return nums;
}

/** Quantos pares consecutivos estão em ordem crescente (mede a ordem de leitura). */
function pontuacaoOrdem(nums: number[]) {
  let s = 0;
  for (let i = 1; i < nums.length; i++) if (nums[i]! > nums[i - 1]!) s++;
  return s;
}

/**
 * Elementos que o esquema OOXML aceita como filhos diretos de <w:body>.
 * Qualquer outro (w:tcPr, w:tblPr, w:trPr, w:tblGrid, ...) só pode existir
 * dentro do seu pai; se vazar para o corpo, o Word acusa arquivo corrompido.
 */
const BLOCOS_DE_CORPO = new Set([
  "w:p",
  "w:tbl",
  "w:sdt",
  "w:customXml",
  "w:bookmarkStart",
  "w:bookmarkEnd",
  "w:commentRangeStart",
  "w:commentRangeEnd",
  "w:proofErr",
  "w:permStart",
  "w:permEnd",
  "w:ins",
  "w:del",
  "w:moveFrom",
  "w:moveTo",
  "w:moveFromRangeStart",
  "w:moveFromRangeEnd",
  "w:moveToRangeStart",
  "w:moveToRangeEnd",
  "w:altChunk",
  "w:sectPr",
]);

/** true quando o elemento pode ficar diretamente dentro de <w:body>. */
export function ehBlocoDeCorpo(el: string) {
  return BLOCOS_DE_CORPO.has(nomeTag(el));
}

/**
 * Percorre recursivamente um elemento do corpo e devolve os elementos
 * "analisáveis" na ordem lógica de leitura.
 *
 * - parágrafo comum -> ele mesmo (ou os parágrafos internos, se for caixa de texto);

 * - tabela usada como layout (contém questões) -> parágrafos das células, na
 *   ordem de leitura correta (linhas ou colunas, decidida pela numeração);
 * - tabela de conteúdo (sem questões) -> preservada inteira.
 */
/** Envelopes que só agrupam blocos (não são conteúdo por si mesmos). */
const ENVELOPES = new Set([
  "w:sdt",
  "w:customXml",
  "w:ins",
  "w:del",
  "w:moveFrom",
  "w:moveTo",
]);

function achatarElemento(el: string): string[] {
  const nome = nomeTag(el);
  if (nome === "w:p") return achatarParagrafo(el);

  if (ENVELOPES.has(nome)) {
    // w:sdt guarda o conteúdo em w:sdtContent; os outros guardam direto
    const conteudos = dividirElementos(conteudoInterno(el)).flatMap((f) =>
      nomeTag(f) === "w:sdtContent" ? dividirElementos(conteudoInterno(f)) : [f],
    );
    const achatados = conteudos.flatMap(achatarElemento).filter(ehBlocoDeCorpo);
    return achatados.length ? achatados : [el];
  }


  if (nome !== "w:tbl") return [el];

  const linhas = dividirElementos(conteudoInterno(el)).filter((x) => nomeTag(x) === "w:tr");
  if (!linhas.length) return [el];

  // grade[linha][celula] = elementos já achatados daquela célula
  const grade: string[][][] = linhas.map((tr) =>
    dividirElementos(conteudoInterno(tr))
      .filter((x) => nomeTag(x) === "w:tc")
      .map((tc) =>
        dividirElementos(conteudoInterno(tc)).flatMap(achatarElemento).filter(ehBlocoDeCorpo),
      ),
  );

  const porLinhas: string[] = grade.flatMap((linha) => linha.flat());
  // tabela sem nenhuma questão: é conteúdo (dados/apoio) — preserva intacta
  if (numerosDeQuestao(porLinhas).length === 0) return [el];

  const colunas = Math.max(...grade.map((l) => l.length));
  const porColunas: string[] = [];
  for (let c = 0; c < colunas; c++)
    for (const linha of grade) if (linha[c]) porColunas.push(...linha[c]!);

  // escolhe a travessia cuja numeração é mais coerente (não presume 1x2)
  const ordemLinhas = pontuacaoOrdem(numerosDeQuestao(porLinhas));
  const ordemColunas = pontuacaoOrdem(numerosDeQuestao(porColunas));
  return ordemColunas > ordemLinhas ? porColunas : porLinhas;
}

/** Caixas de texto: expande os parágrafos internos quando contêm questões. */
function achatarParagrafo(p: string): string[] {
  if (!/<w:txbxContent\b/.test(p)) return [p];
  // mc:Fallback repete o mesmo conteúdo da caixa de texto: ignorar para não duplicar
  const limpo = p.replace(/<mc:Fallback(?:\s[^>]*)?>[\s\S]*?<\/mc:Fallback>/g, "");
  const internos: string[] = [];
  const re = /<w:txbxContent(?:\s[^>]*)?>([\s\S]*?)<\/w:txbxContent>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(limpo))) {
    internos.push(...dividirElementos(m[1]!).flatMap(achatarElemento).filter(ehBlocoDeCorpo));
  }
  if (!numerosDeQuestao(internos).length) return [p];
  return internos;
}



/* ------------------------------------------------------------------ */
/* 2. Question Detector                                                */
/* ------------------------------------------------------------------ */

/**
 * Aceita: "1.", "1)", "01.", "01)", "1 -", "1 –", "Questão 1", "QUESTÃO 1:",
 * "Questão 01 -", "Q1)", "(1)" e variações de espaço/caixa.
 */
const RE_QUESTAO_ROTULADA =
  /^\s*(?:quest(?:[ãa]o|\.)|q)\s*[nº°.:-]*\s*\(?\s*(\d{1,3})\s*\)?\s*(?:[.)\-–—:]\s*|\s+|$)/i;
const RE_QUESTAO_NUMERO = /^\s*\(?\s*(\d{1,3})\s*\)?\s*([.)\-–—:])\s*/;
/** "(1) enunciado", "( 01 ) enunciado" — parênteses sem pontuação extra */
const RE_QUESTAO_PARENTESES = /^\s*\(\s*(\d{1,3})\s*\)\s*/;

const RE_ALTERNATIVA = /^\s*\(?\s*([A-Za-z])\s*[).\-–—:]\s+?/;
const RE_AFIRMACAO = /^\s*[([]\s*[VvFfXx]?\s*[)\]]\s*/;
const RE_ROMANO = /^\s*\(?\s*(?:[IVXivx]{1,5})\s*[).\-–—:]\s+/;

/** Frases que costumam abrir/indicar um texto de apoio compartilhado. */
const RE_ANCORA =
  /(leia\s+(o|a|os|as|atentamente)\b|leia\s+o\s+texto|texto\s+(base|de\s+apoio|para\s+(as\s+)?quest)|com\s+base\s+n[oa]s?\s+|de\s+acordo\s+com\s+[oa]s?\s+(texto|poema|trecho|imagem|tabela|gr[áa]fico|charge|tirinha)|segundo\s+[oa]\s+(texto|autor|poema|trecho)|ap[óo]s\s+a\s+leitura|analise\s+[oa]s?\s+(texto|imagem|tabela|gr[áa]fico|charge|tirinha|trecho)|observe\s+[oa]s?\s+(imagem|figura|tabela|gr[áa]fico|tirinha|charge)|leitura\s+do\s+texto)/i;

/**
 * Parágrafo/célula que contém APENAS o marcador da questão ("1.", "2)", "(3)",
 * "Questão 4", "4"). Muito comum em provas montadas em tabela, onde o número
 * fica numa célula e o enunciado na célula ao lado.
 */
const RE_SO_MARCADOR = /^\s*(?:quest(?:[ãa]o|\.)|q)?\s*[nº°.:-]*\s*\(?\s*(\d{1,3})\s*\)?\s*[.)\-–—:]?\s*$/i;

export type MarcaQuestao = {
  numero: number;
  rotulo: string;
  soMarcador: boolean;
  /** marcador "fraco": só dígitos, sem rótulo nem pontuação ("18" numa célula
   *  de tabela de dados). Não serve para decidir se uma tabela é de layout. */
  fraco: boolean;
};

function detectarQuestao(texto: string): MarcaQuestao | null {
  const limpo = texto.trim();
  if (!limpo) return null;
  if (RE_AFIRMACAO.test(limpo)) return null;

  const rotulada = RE_QUESTAO_ROTULADA.exec(limpo);
  if (rotulada) {
    const numero = Number(rotulada[1]);
    if (numero >= 1 && numero <= 300)
      return {
        numero,
        rotulo: rotulada[0],
        soMarcador: !limpo.slice(rotulada[0].length).trim(),
        fraco: false,
      };
  }

  const simples = RE_QUESTAO_NUMERO.exec(limpo);
  if (simples) {
    const numero = Number(simples[1]);
    if (numero < 1 || numero > 300) return null;
    const resto = limpo.slice(simples[0].length).trim();
    if (resto) return { numero, rotulo: simples[0], soMarcador: false, fraco: false };
  }

  const parenteses = RE_QUESTAO_PARENTESES.exec(limpo);
  if (parenteses) {
    const numero = Number(parenteses[1]);
    if (numero < 1 || numero > 300) return null;
    if (limpo.slice(parenteses[0].length).trim())
      return { numero, rotulo: parenteses[0], soMarcador: false, fraco: false };
  }

  // só o marcador, sem enunciado no mesmo parágrafo/célula
  const so = limpo.length <= 14 ? RE_SO_MARCADOR.exec(limpo) : null;
  if (so) {
    const numero = Number(so[1]);
    if (numero >= 1 && numero <= 300)
      return {
        numero,
        rotulo: limpo,
        soMarcador: true,
        // "18" puro (sem "Questão", sem ".", ")" etc.) é ambíguo: pode ser um
        // dado dentro de uma tabela de conteúdo.
        fraco: /^\d{1,3}$/.test(limpo),
      };
  }
  return null;
}

type Cand = { idx: number; numero: number; rotulo: string; rotulada: boolean; auto: boolean };

/**
 * Escolhe a maior subsequência crescente de candidatos (mantendo a ordem do
 * documento). Isso tolera numeração que começa em 4, saltos e falsos
 * positivos isolados (datas, valores, itens de lista).
 */
function melhorSequencia(cands: Cand[]): Cand[] {
  if (!cands.length) return [];
  const melhor: number[] = new Array(cands.length).fill(1);
  const anterior: number[] = new Array(cands.length).fill(-1);
  const peso = (c: Cand) => (c.rotulada ? 2 : 1);
  const score: number[] = cands.map((c) => peso(c));

  for (let i = 0; i < cands.length; i++) {
    for (let j = 0; j < i; j++) {
      // numeração manual e automática podem coexistir e repetir um número
      // (duas contagens diferentes); nesse caso aceitamos empate.
      // Também aceita reinício explícito da numeração (ex.: 1..10 de Inglês,
      // depois 1..10 de Espanhol). Antes, a subsequência crescente descartava
      // toda disciplina que reiniciasse em 1 e seu conteúdo acabava anexado à
      // última questão da disciplina anterior.
      const reinicioDeSecao = cands[i]!.numero === 1 && cands[j]!.numero > 1;
      const compativel =
        cands[j]!.numero < cands[i]!.numero ||
        reinicioDeSecao ||
        (cands[j]!.auto !== cands[i]!.auto && cands[j]!.numero <= cands[i]!.numero);
      if (compativel && score[j]! + peso(cands[i]!) > score[i]!) {
        score[i] = score[j]! + peso(cands[i]!);
        melhor[i] = melhor[j]! + 1;
        anterior[i] = j;
      }
    }
  }
  let fim = 0;
  for (let i = 1; i < cands.length; i++) if (score[i]! > score[fim]!) fim = i;
  const saida: Cand[] = [];
  for (let i = fim; i >= 0; i = anterior[i]!) {
    saida.push(cands[i]!);
    if (anterior[i]! < 0) break;
  }
  return saida.reverse();
}

/* ------------------------------------------------------------------ */
/* Pipeline principal                                                  */
/* ------------------------------------------------------------------ */

export async function analisarDocx(bytes: ArrayBuffer, nomeArquivo = "documento.docx"): Promise<AnaliseProva> {
  const zip = await JSZip.loadAsync(bytes.slice(0));
  const arquivo = zip.file("word/document.xml");
  if (!arquivo) throw new Error("Arquivo DOCX inválido: documento principal não encontrado.");
  const xml = await arquivo.async("string");
  const numberingXml = (await zip.file("word/numbering.xml")?.async("string")) ?? "";
  const stylesXml = (await zip.file("word/styles.xml")?.async("string")) ?? "";
  const mapaNumeracao = lerMapaNumeracao(numberingXml, stylesXml);
  const { prefixo, sufixo, blocos: brutos } = dividirCorpo(xml);

  let sectPr = "";
  const blocos: Bloco[] = [];
  for (const bruto of brutos) {
    if (bruto.startsWith("<w:sectPr")) {
      sectPr = bruto;
      continue;
    }
    // percorre toda a estrutura (tabelas, linhas, células, caixas de texto)
    for (const b of achatarElemento(bruto)) {
      // segurança: nunca deixa um elemento inválido no nível de <w:body>
      if (!ehBlocoDeCorpo(b)) continue;
      const nome = nomeTag(b);

      blocos.push({
        xml: b,
        texto: textoDoBloco(b),
        tipo: nome === "w:p" ? "p" : nome === "w:tbl" ? "tbl" : "outro",
        temImagem: /<w:drawing\b|<w:pict\b|<v:imagedata\b/.test(b),
        num: null,
      });
    }
  }


  /* --- numeração automática do Word (w:numPr/w:numId/w:ilvl) ---------- */
  const marcas = calcularNumeracao(
    blocos.map((b, indice) => ({ indice, xml: b.xml, ehParagrafo: b.tipo === "p" })),
    mapaNumeracao,
  );
  marcas.forEach((marca, idx) => {
    blocos[idx]!.num = marca;
  });

  /* --- candidatos a início de questão --------------------------------- */
  /* Regra: primeiro a numeração escrita no texto; só quando ela não existe
     é que a numeração automática do Word entra como candidata.            */
  const candidatos: Cand[] = [];
  let candidatosAuto = 0;
  blocos.forEach((b, idx) => {
    if (b.tipo !== "p") return;
    const q = detectarQuestao(b.texto);
    if (q) {
      // marcador isolado ("1." numa célula) só vale se houver conteúdo depois
      if (q.soMarcador) {
        const seguinte = blocos
          .slice(idx + 1, idx + 5)
          .some((s) => s.temImagem || s.texto.trim().length > 3);
        if (!seguinte) return;
      }
      candidatos.push({
        idx,
        numero: q.numero,
        rotulo: q.rotulo,
        rotulada: /quest/i.test(q.rotulo),
        auto: false,
      });
      return;
    }

    const marca = b.num;
    if (!marca || marca.ilvl !== 0 || marca.formato !== "decimal") return;
    if (!b.texto.trim()) return;
    candidatosAuto++;
    candidatos.push({
      idx,
      numero: marca.valor,
      rotulo: "",
      rotulada: false,
      auto: true,
    });
  });

  const inicios = melhorSequencia(candidatos);

  let cabecalho = inicios.length
    ? Array.from({ length: inicios[0]!.idx }, (_, i) => i)
    : blocos.map((_, i) => i);

  const questoes: Questao[] = inicios.map((ini, i) => {
    const fim = i + 1 < inicios.length ? inicios[i + 1]!.idx : blocos.length;
    const indices = Array.from({ length: fim - ini.idx }, (_, k) => ini.idx + k);
    return montarQuestao(blocos, indices, ini.numero, ini.rotulo, i, ini.auto);
  });

  /* --- 3. Support Text Detector --------------------------------------- */
  /* Blocos que sobram depois da última alternativa de uma questão e antes da
     próxima questão são o texto de apoio da próxima — a menos que sejam só
     linhas em branco ou linhas de resposta da própria questão anterior.     */
  const apoios: number[][] = questoes.map(() => []);
  const vazio = (i: number) => !blocos[i]!.texto.trim() && !blocos[i]!.temImagem;
  const soLinhas = (i: number) => /^[\s_.\-–—]+$/.test(blocos[i]!.texto);

  for (let i = 0; i < questoes.length - 1; i++) {
    const q = questoes[i]!;
    if (!q.rodape.length) continue;
    const util = q.rodape.filter((b) => !vazio(b) && !soLinhas(b));
    if (!util.length) continue;

    // ponto de corte: primeiro bloco "útil" do rodapé; o que vem antes
    // (brancos/linhas de resposta) continua pertencendo à questão anterior.
    const corte = q.rodape.indexOf(util[0]!);
    apoios[i + 1] = q.rodape.slice(corte);
    q.rodape = q.rodape.slice(0, corte);
  }

  // texto de apoio da primeira questão, quando está junto do cabeçalho
  if (questoes.length && cabecalho.length > 1) {
    let pos = cabecalho.findIndex((i) => RE_ANCORA.test(blocos[i]!.texto));
    if (pos < 0) {
      // heurística: parágrafo longo / imagem / tabela imediatamente antes da
      // primeira questão, precedido por um título curto
      const ultimos = cabecalho.slice(-6);
      const achou = ultimos.findIndex(
        (i) => blocos[i]!.texto.trim().length > 180 || blocos[i]!.temImagem || blocos[i]!.tipo === "tbl",
      );
      if (achou >= 0) pos = cabecalho.length - ultimos.length + Math.max(0, achou - 1);
    }
    if (pos > 0) {
      apoios[0] = cabecalho.slice(pos);
      cabecalho = cabecalho.slice(0, pos);
    }
  }

  /* --- 4. Question Grouping ------------------------------------------- */
  const contextos: Contexto[] = [];
  let letraTexto = 0;
  let contadorIndep = 0;
  questoes.forEach((q, i) => {
    const abreNovo = !contextos.length || apoios[i]!.length > 0;
    if (abreNovo) {
      const temApoio = apoios[i]!.length > 0;
      const id = temApoio
        ? `texto_${String.fromCharCode(65 + letraTexto++)}`
        : `independentes_${++contadorIndep}`;
      contextos.push({ id, apoio: apoios[i]!, questoes: [i] });
    } else {
      contextos[contextos.length - 1]!.questoes.push(i);
    }
    const ctx = contextos[contextos.length - 1]!;
    q.supportId = ctx.apoio.length ? ctx.id : null;
  });

  const diagnostico: Diagnostico = {
    blocos: blocos.length,
    candidatos: candidatos.length,
    candidatosAuto,
    questoes: questoes.length,
    textosApoio: contextos.filter((c) => c.apoio.length).length,
    independentes: questoes.filter((q) => !q.supportId).length,
    imagens: blocos.filter((b) => b.temImagem).length,
    tabelas: blocos.filter((b) => b.tipo === "tbl").length,
    grupos: contextos.map((c) => ({
      id: c.id,
      questoes: c.questoes.map((i) => questoes[i]!.id),
    })),
  };

  console.info(`[EmbaralhaProvas] Arquivo: ${nomeArquivo}`);
  console.info(`[EmbaralhaProvas] Questões encontradas: ${diagnostico.questoes}`);
  console.info(`[EmbaralhaProvas] Textos de apoio encontrados: ${diagnostico.textosApoio}`);
  console.info(`[EmbaralhaProvas] Questões independentes: ${diagnostico.independentes}`);
  console.info(`[EmbaralhaProvas] Imagens encontradas: ${diagnostico.imagens}`);
  console.info(`[EmbaralhaProvas] Tabelas encontradas: ${diagnostico.tabelas}`);
  diagnostico.grupos.forEach((g) => console.info(`[EmbaralhaProvas] ${g.id}: ${g.questoes.join(", ")}`));

  if (!diagnostico.questoes) {
    console.warn("[DOCX DEBUG]");
    console.warn(`blocos encontrados: ${diagnostico.blocos}`);
    console.warn(`questões detectadas: ${diagnostico.questoes}`);
    console.warn(`candidatos: ${diagnostico.candidatos}`);
    console.warn(`questões independentes: ${diagnostico.independentes}`);
    console.warn(
      "[DOCX DEBUG] primeiros parágrafos:",
      blocos.slice(0, 15).map((b) => b.texto.slice(0, 80)),
    );
  }


  return { blocos, cabecalho, questoes, contextos, sectPr, prefixo, sufixo, bytes, diagnostico };
}

/* ------------------------------------------------------------------ */
/* QuestionBlock                                                       */
/* ------------------------------------------------------------------ */

function letraSeguinte(l: string) {
  return String.fromCharCode(l.toUpperCase().charCodeAt(0) + 1);
}

function montarQuestao(
  blocos: Bloco[],
  indices: number[],
  numeroOriginal: number,
  rotuloOriginal: string,
  posicaoOriginal: number,
  numeracaoAuto = false,
): Questao {
  /* --- marcadores de alternativas ------------------------------------ */
  type Marca = { idx: number; letra: string; rotulo: string };
  const marcas: Marca[] = [];
  for (const idx of indices.slice(1)) {
    const b = blocos[idx]!;
    if (b.tipo !== "p") continue;
    if (RE_ROMANO.test(b.texto)) continue; // itens I/II/III fazem parte do enunciado
    const m = RE_ALTERNATIVA.exec(b.texto);
    if (m && b.texto.slice(m[0].length).trim()) {
      marcas.push({ idx, letra: m[1]!, rotulo: m[0] });
      continue;
    }
    // alternativas numeradas automaticamente pelo Word (a), b), c) ...)
    const auto = b.num;
    if (auto && auto.formato === "letra" && b.texto.trim()) {
      marcas.push({
        idx,
        letra: String.fromCharCode(97 + ((auto.valor - 1) % 26)),
        rotulo: "",
      });
    }
  }

  // sequência válida: começa em A/a e segue consecutiva (A..Z)
  let sequencia: Marca[] = [];
  for (let i = 0; i < marcas.length; i++) {
    if (marcas[i]!.letra.toUpperCase() !== "A") continue;
    const atual: Marca[] = [marcas[i]!];
    for (let j = i + 1; j < marcas.length; j++) {
      const esperada = letraSeguinte(atual[atual.length - 1]!.letra);
      if (marcas[j]!.letra.toUpperCase() === esperada) atual.push(marcas[j]!);
    }
    if (atual.length > sequencia.length) sequencia = atual;
    break;
  }

  const alternativas: Alternativa[] = [];
  let rodape: number[] = [];
  let corpo: number[] = indices.slice();

  if (sequencia.length >= 2) {
    const primeiro = sequencia[0]!.idx;
    corpo = indices.filter((i) => i < primeiro);
    const ultimoIdx = sequencia[sequencia.length - 1]!.idx;
    sequencia.forEach((marca, k) => {
      const proximo = k + 1 < sequencia.length ? sequencia[k + 1]!.idx : ultimoIdx + 1;
      alternativas.push({
        letra: marca.letra,
        rotulo: marca.rotulo,
        indices: indices.filter((i) => i >= marca.idx && i < proximo),
      });
    });
    rodape = indices.filter((i) => i > ultimoIdx);
  }

  /* --- afirmações V/F (apenas quando não há alternativas) ------------- */
  const afirmacoes: Afirmacao[] = [];
  if (!alternativas.length) {
    const marcasVf = corpo
      .slice(1)
      .filter((idx) => blocos[idx]!.tipo === "p" && RE_AFIRMACAO.test(blocos[idx]!.texto));
    if (marcasVf.length >= 2) {
      const primeiro = marcasVf[0]!;
      const ultimo = marcasVf[marcasVf.length - 1]!;
      marcasVf.forEach((idx, k) => {
        const proximo = k + 1 < marcasVf.length ? marcasVf[k + 1]! : ultimo + 1;
        afirmacoes.push({
          rotulo: RE_AFIRMACAO.exec(blocos[idx]!.texto)![0],
          indices: corpo.filter((i) => i >= idx && i < proximo),
        });
      });
      rodape = corpo.filter((i) => i > ultimo).concat(rodape);
      corpo = corpo.filter((i) => i < primeiro);
    }
  }

  const primeiroTexto = blocos[indices[0]!]!.texto;
  const semRotulo = rotuloOriginal && primeiroTexto.trimStart().startsWith(rotuloOriginal.trim())
    ? primeiroTexto.trimStart().slice(rotuloOriginal.trim().length)
    : primeiroTexto;
  // quando o marcador está isolado (número numa célula), o enunciado vem depois
  const textoPreview =
    semRotulo.trim() ||
    (corpo.length ? corpo.map((i) => blocos[i]!.texto.trim()).find((t) => t.length > 3) : "") ||
    indices.map((i) => blocos[i]!.texto.trim()).find((t) => t.length > 3) ||
    "";
  const preview = textoPreview.trim().slice(0, 140) || "(sem texto)";


  return {
    id: `q${numeroOriginal}_${posicaoOriginal}`,
    supportId: null,
    numeroOriginal,
    posicaoOriginal,
    rotuloOriginal,
    numeracaoAuto,
    preview,
    corpo,
    alternativas,
    afirmacoes,
    rodape,
  };
}
