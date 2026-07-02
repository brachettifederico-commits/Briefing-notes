import os
import json
import re
import io
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template
import anthropic
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
MODEL = "claude-sonnet-5"

SEARCH_PROMPT = """Cerca informazioni aggiornate e dettagliate su questa persona:
- Nome: {nome}{istituzione}
- Focus tematico di interesse: {focus}

Trova e riporta in modo dettagliato:
1. Data e luogo di nascita
2. Ruolo istituzionale attuale
3. Carriera professionale completa (cronologica)
4. Carriera politica completa (cronologica)
5. Almeno 4-6 posizioni, dichiarazioni o attività SPECIFICHE e VERIFICABILI sul tema "{focus}" — non punti generici. Per ognuna cerca una fonte diversa quando possibile (articoli di stampa, comunicati istituzionali, atti parlamentari, interviste).

Regole per le fonti:
- Riporta l'URL ESATTO della pagina che hai effettivamente consultato per ogni fatto, subito dopo il fatto stesso.
- Non riutilizzare lo stesso URL per più di 2 punti diversi: cerca fonti distinte per ciascuna informazione sul focus tematico.
- Se non trovi una fonte affidabile per un punto, ometti il punto piuttosto che inventare una fonte.
- Preferisci fonti primarie (siti istituzionali, comunicati ufficiali, atti parlamentari) e testate giornalistiche riconosciute.

Sii preciso e dettagliato. Includi date, ruoli e contesti specifici."""

JSON_SYSTEM = """Sei un assistente che redige briefing note istituzionali italiane per uno studio di public affairs.

Ti viene fornito un testo con informazioni su una persona, già corredato di URL fonte per ciascun fatto. Struttura queste informazioni in JSON valido.

Regole assolute:
- Rispondi SOLO con JSON valido, nessun testo prima o dopo
- Nessun tag HTML, nessun markdown, nessun <cite>, testo plain
- Italiano formale e istituzionale
- Se un'informazione non è disponibile scrivi "Dato non disponibile"
- Per ogni campo "_url" e ogni "url" nei focus_items: riporta l'URL esatto fornito nel testo di ricerca per quel punto specifico. Se il testo non fornisce un URL per un dato punto, usa null — non inventare e non riutilizzare un URL di un altro punto.
- Nei focus_items, preferisci URL diversi tra loro quando il testo di ricerca li fornisce distinti.

Formato JSON esatto:
{
  "nome": "Nome Cognome",
  "luogo_nascita": "Nato/a a Città (Provincia) il GG mese AAAA",
  "ruolo_attuale": "Ruolo istituzionale attuale completo",
  "background_sintetico": "Ex X, Y e Z (max 10 parole)",
  "istituzione_contesto": "Es: Camera dei Deputati",
  "carriera_professionale": "Testo narrativo prosa cronologico, 150-250 parole, stile istituzionale.",
  "carriera_professionale_url": "URL fonte oppure null",
  "carriera_politica": "Testo narrativo prosa cronologico, 150-250 parole, stile istituzionale.",
  "carriera_politica_url": "URL fonte oppure null",
  "focus_titolo": "Titolo focus tematico (es: Digitale, IA e Cybersecurity)",
  "focus_items": [
    {
      "titolo": "Titolo breve del punto",
      "testo": "Descrizione della posizione o attività. 2-4 frasi plain text.",
      "url": "URL fonte oppure null"
    }
  ]
}

focus_items: 4-6 voci rilevanti al focus tematico richiesto, ciascuna con fonte propria quando disponibile."""


def strip_cite(text):
    """Rimuove tag <cite> e markup dal testo."""
    if not text:
        return text
    text = re.sub(r'<cite[^>]*>([\s\S]*?)</cite>', r'\1', text)
    text = re.sub(r'</?cite[^>]*>', '', text)
    text = re.sub(r'\[?\(Link\)\]?\([^)]*\)', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def clean_obj(obj):
    """Pulisce ricorsivamente tutti i valori stringa da tag HTML."""
    if isinstance(obj, str):
        return strip_cite(obj)
    elif isinstance(obj, list):
        return [clean_obj(i) for i in obj]
    elif isinstance(obj, dict):
        return {k: clean_obj(v) for k, v in obj.items()}
    return obj


def call_anthropic_with_search(prompt):
    """Chiama Anthropic con web search e gestisce il loop tool_use."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    messages = [{"role": "user", "content": prompt}]
    tools = [{"type": "web_search_20250305", "name": "web_search"}]
    research_text = ""

    for _ in range(10):
        response = client.messages.create(
            model=MODEL,
            max_tokens=4000,
            tools=tools,
            messages=messages,
        )

        messages.append({"role": "assistant", "content": response.content})

        if response.stop_reason == "end_turn":
            for block in response.content:
                if hasattr(block, "text"):
                    research_text = block.text
            break

        if response.stop_reason == "tool_use":
            tool_results = []
            for block in response.content:
                if block.type == "tool_use":
                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": "Continua la ricerca.",
                    })
            messages.append({"role": "user", "content": tool_results})

    return research_text


def call_anthropic_json(research_text, focus, data_meeting):
    """Seconda chiamata senza web search per formattare in JSON pulito."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    user_msg = f"""Sulla base di queste informazioni, genera la briefing note in JSON:

INFORMAZIONI RACCOLTE:
{research_text}

FOCUS TEMATICO PER IL CLIENTE: {focus}
{f'DATA MEETING: {data_meeting}' if data_meeting else ''}

Rispondi SOLO con JSON valido, testo plain senza tag HTML."""

    response = client.messages.create(
        model=MODEL,
        max_tokens=4000,
        system=JSON_SYSTEM,
        messages=[{"role": "user", "content": user_msg}],
    )

    for block in response.content:
        if hasattr(block, "text"):
            return block.text

    return None


def add_hyperlink(paragraph, text, url):
    """Aggiunge un hyperlink cliccabile a un paragrafo Word."""
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rPr.append(rStyle)

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "1155CC")
    rPr.append(color)

    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)

    new_run.append(rPr)

    t = OxmlElement("w:t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    new_run.append(t)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def set_cell_border(cell, color="CCCCCC"):
    """Imposta i bordi di una cella."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ["top", "left", "bottom", "right"]:
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:color"), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)


def set_cell_shading(cell, fill="F2F2F2"):
    """Imposta il colore di sfondo di una cella."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def generate_docx(d, data_meeting):
    """Genera il documento Word nel formato CZP."""
    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)

    styles = doc.styles
    try:
        hl_style = styles["Hyperlink"]
    except KeyError:
        hl_style = styles.add_style("Hyperlink", 2)
        hl_style.font.color.rgb = RGBColor(0x11, 0x55, 0xCC)
        hl_style.font.underline = True

    if data_meeting:
        try:
            dt = datetime.strptime(data_meeting, "%Y-%m-%d")
            today = dt.strftime("%-d %B %Y")
        except Exception:
            today = data_meeting
    else:
        today = datetime.now().strftime("%-d %B %Y")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("BRIEFING NOTE")
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    run.font.name = "Arial"

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(d.get("istituzione_contesto", ""))
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    run.font.name = "Arial"

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(today)
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x77, 0x77, 0x77)
    run.font.name = "Arial"

    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run(d.get("nome", ""))
    run.bold = True
    run.font.size = Pt(26)
    run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    run.font.name = "Arial"

    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "1F3864")
    pBdr.append(bottom)
    pPr.append(pBdr)

    for text in [d.get("luogo_nascita", ""), d.get("ruolo_attuale", ""), d.get("background_sintetico", "")]:
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.size = Pt(10)
        run.font.name = "Arial"

    doc.add_paragraph()

    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"

    col_widths = [Cm(4.5), Cm(13.5)]

    def add_table_row(label, text, url=None, is_focus=False, focus_items=None):
        row = table.add_row()
        row.cells[0].width = col_widths[0]
        row.cells[1].width = col_widths[1]

        set_cell_border(row.cells[0])
        set_cell_shading(row.cells[0])
        p_label = row.cells[0].paragraphs[0]
        run = p_label.add_run(label)
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = "Arial"

        set_cell_border(row.cells[1])

        if is_focus and focus_items:
            cell = row.cells[1]
            first = True
            for item in focus_items:
                if first:
                    p_content = cell.paragraphs[0]
                    first = False
                else:
                    p_content = cell.add_paragraph()

                run_title = p_content.add_run(item.get("titolo", "") + ". ")
                run_title.bold = True
                run_title.font.size = Pt(10)
                run_title.font.name = "Arial"

                run_text = p_content.add_run(item.get("testo", ""))
                run_text.font.size = Pt(10)
                run_text.font.name = "Arial"

                if item.get("url"):
                    run_space = p_content.add_run(" ")
                    run_space.font.size = Pt(10)
                    add_hyperlink(p_content, "(Link)", item["url"])
        else:
            p_content = row.cells[1].paragraphs[0]
            run_text = p_content.add_run(text or "")
            run_text.font.size = Pt(10)
            run_text.font.name = "Arial"
            if url:
                run_space = p_content.add_run(" ")
                run_space.font.size = Pt(10)
                add_hyperlink(p_content, "(Link)", url)

    add_table_row(
        "Carriera professionale",
        d.get("carriera_professionale", ""),
        d.get("carriera_professionale_url")
    )
    add_table_row(
        "Carriera politica",
        d.get("carriera_politica", ""),
        d.get("carriera_politica_url")
    )
    add_table_row(
        f"Focus: {d.get('focus_titolo', '')}",
        None,
        is_focus=True,
        focus_items=d.get("focus_items", [])
    )

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/generate", methods=["POST"])
def generate():
    data = request.get_json()
    nome = data.get("nome", "").strip()
    istituzione = data.get("istituzione", "").strip()
    focus = data.get("focus", "").strip()
    data_meeting = data.get("dataMeeting", "")

    if not nome or not focus:
        return jsonify({"error": "Nome e focus sono obbligatori."}), 400

    try:
        istituzione_str = f"\n- Istituzione / ruolo: {istituzione}" if istituzione else ""
        search_prompt = SEARCH_PROMPT.format(
            nome=nome,
            istituzione=istituzione_str,
            focus=focus
        )
        research_text = call_anthropic_with_search(search_prompt)

        if not research_text:
            return jsonify({"error": "Nessun risultato dalla ricerca."}), 500

        clean_research = strip_cite(research_text)

        json_text = call_anthropic_json(clean_research, focus, data_meeting)
        if not json_text:
            return jsonify({"error": "Nessun JSON nella risposta."}), 500

        clean_json = re.sub(r'```json|```', '', json_text).strip()
        parsed = json.loads(clean_json)
        parsed = clean_obj(parsed)

        return jsonify(parsed)

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/docx", methods=["POST"])
def docx_endpoint():
    data = request.get_json()
    data_meeting = data.get("dataMeeting", "")

    try:
        buf = generate_docx(data, data_meeting)
        nome = data.get("nome", "briefing").replace(" ", "_")
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=f"Briefing_Note_{nome}.docx"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
