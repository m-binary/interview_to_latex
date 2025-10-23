import streamlit as st
import pandas as pd
from io import BytesIO


TEX_BODY = r"""
\section*{{Expertengespräch mit {name}}}

\begin{{minipage}}{{\textwidth}}
		\begin{{tabbing}}
		Wissenschaftliche Betreuerin: \hspace{{0.85cm}}\=\kill
		Interviewpartner: \> {interviewpartner}\\[1mm]
		Funktion: \> {funktion} \\[3mm]
		Interviewer: \> {interviewer} \\[1mm]
		Gesprächszeitpunkt: \> {zeitpunkt}\\[1mm]
		Ort: \> {ort}\\[3mm]
		Thema: \> {thema}\\

		\end{{tabbing}}
	\end{{minipage}}

\begin{{longtable}}{{|p{{2cm}}|p{{13cm}}|}}
\hline
{rows}
\hline
\end{{longtable}}
"""

TEX_PREAMBLE = r"""
\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[ngerman]{babel}
\usepackage{longtable}
\usepackage{geometry}
\usepackage{ragged2e}
\usepackage{array}
\usepackage{booktabs}
\usepackage{hyperref}
\geometry{left=2.5cm,right=2.5cm,top=2.5cm,bottom=2.5cm}

\begin{document}
"""

TEX_SIGNATURE_PAGE = r"""
\newpage
\section*{{Bestätigung der Gesprächsinhalte}}

\vspace{{1cm}}

Hiermit bestätigen die Unterzeichnenden, dass die Inhalte des vorliegenden Expertengesprächs korrekt wiedergegeben wurden und für die Verwendung in der Bachelorarbeit freigegeben werden.

\vspace{{2cm}}

\noindent
\begin{{minipage}}{{0.45\textwidth}}
\centering
\rule{{0.8\linewidth}}{{0.4pt}}\\
Ort, Datum
\end{{minipage}}
\hfill
\begin{{minipage}}{{0.45\textwidth}}
\centering
\rule{{0.8\linewidth}}{{0.4pt}}\\
{interviewpartner}\\
Interviewpartner
\end{{minipage}}

\vspace{{2cm}}

\noindent
\begin{{minipage}}{{0.45\textwidth}}
\centering
\rule{{0.8\linewidth}}{{0.4pt}}\\
Ort, Datum
\end{{minipage}}
\hfill
\begin{{minipage}}{{0.45\textwidth}}
\centering
\rule{{0.8\linewidth}}{{0.4pt}}\\
{interviewer}\\
Interviewer
\end{{minipage}}
"""

TEX_END = r"""

\end{document}
"""


def latex_escape(s: str) -> str:
    """Escape common LaTeX special characters in a string."""
    if s is None:
        return ""
    s = str(s)
    replacements = {
        '\\': r'\\textbackslash{}',
        '&': r'\\&',
        '%': r'\\%',
        '$': r'\\$',
        '#': r'\\#',
        '_': r'\\_',
        '{': r'\\{',
        '}': r'\\}',
        '~': r'\\textasciitilde{}',
        '^': r'\\^{}',
        '€': r'\\euro{}',
    }
    for k, v in replacements.items():
        s = s.replace(k, v)
    # Collapse multiple whitespace/newlines into single space
    s = ' '.join(s.split())
    return s


def extract_last_name(full_name: str) -> str:
    """Extract the last name from a full name."""
    if not full_name or not full_name.strip():
        return full_name
    parts = full_name.strip().split()
    return parts[-1] if parts else full_name


def df_to_latex_rows(df: pd.DataFrame, use_last_name_only: bool = False) -> str:
    """Convert two-column dataframe (speaker, text) to LaTeX longtable rows with escaping."""
    lines = []
    for _, row in df.iterrows():
        speaker_raw = str(row.iloc[0])
        if use_last_name_only:
            speaker_raw = extract_last_name(speaker_raw)
        speaker = latex_escape(speaker_raw)
        text = latex_escape(row.iloc[1])
        lines.append(f"{speaker} & {text}\\\\")
    return "\n".join(lines)


def generate_tex(metadata: dict, df: pd.DataFrame, use_last_name_only: bool = False, include_signature: bool = True) -> str:
    rows = df_to_latex_rows(df, use_last_name_only)
    body = TEX_BODY.format(rows=rows, **metadata)
    
    signature = ""
    if include_signature:
        signature = TEX_SIGNATURE_PAGE.format(**metadata)
    
    full = TEX_PREAMBLE + body + signature + TEX_END
    return full


def app():
    st.title("Interview → LaTeX Konverter")
    st.markdown("Lade eine Excel-Datei (.xlsx) mit zwei Spalten (ohne Kopfzeile): Teilnehmer | Inhalt")

    # Metadata fields extracted from example
    with st.expander("Metadaten (werden in die LaTeX-Vorlage eingefügt)"):
        name = st.text_input("Titel / Name (z.B. 'Expertengespräch mit Max Mustermann')", value="")
        interviewpartner = st.text_input("Interviewpartner", value="")
        funktion = st.text_input("Funktion", value="")
        interviewer = st.text_input("Interviewer", value="")
        zeitpunkt = st.text_input("Gesprächszeitpunkt", value="")
        ort = st.text_input("Ort", value="")
        thema = st.text_input("Thema", value="")

    # Display options
    st.subheader("Anzeigeoptionen")
    use_last_name_only = st.checkbox("Nur Nachnamen in der Tabelle anzeigen", value=False, 
                                      help="Zeigt nur den Nachnamen der Sprecher in der Interview-Tabelle an")
    include_signature = st.checkbox("Unterschriftenseite einfügen", value=True,
                                     help="Fügt eine Bestätigungsseite am Ende des Dokuments ein")

    uploaded = st.file_uploader("Excel-Datei (.xlsx)", type=["xlsx"])

    if uploaded is not None:
        try:
            df = pd.read_excel(uploaded, header=None)
            if df.shape[1] < 2:
                st.error("Die Excel-Datei muss mindestens zwei Spalten enthalten (Teilnehmer, Inhalt)")
                return
            st.success(f"Geladen: {df.shape[0]} Zeilen")
            st.dataframe(df)

            # Preview LaTeX rows
            st.subheader("Vorschau LaTeX-Inhalt")
            st.code(df_to_latex_rows(df, use_last_name_only))

            if st.button("Kompilieren"):
                metadata = {
                    "name": name or interviewpartner,
                    "interviewpartner": interviewpartner or name,
                    "funktion": funktion,
                    "interviewer": interviewer,
                    "zeitpunkt": zeitpunkt,
                    "ort": ort,
                    "thema": thema,
                }
                tex = generate_tex(metadata, df, use_last_name_only, include_signature)
                b = tex.encode("utf-8")
                st.download_button("LaTeX-Datei herunterladen", data=b, file_name="interview.tex", mime="text/x-tex")

        except Exception as e:
            st.error(f"Fehler beim Einlesen der Datei: {e}")
    else:
        st.info("Bitte lade eine Excel-Datei hoch.")


if __name__ == "__main__":
    app()
