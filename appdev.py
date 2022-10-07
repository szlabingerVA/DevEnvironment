'''
Voestalpine Signaling Analyse-App
'''
import io
from itertools import cycle
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from streamlit_option_menu import option_menu
from streamlit_echarts import st_echarts
from st_aggrid import AgGrid, GridOptionsBuilder

# Webapp Konfiguration
st.set_page_config(page_title="Analyse Dashboard",
                   page_icon=":bar_chart:", layout="centered")

# Farbpaletten definieren
px.colors.qualitative.VoestGrey = ["#E3E3E3", "#C4C4C4", "#A5A5A5"]
px.colors.qualitative.VoestBlue = ["#91C8DC", "#50AACD", "#0082B4"]

# Chartfarbpalette
palette = cycle(["#A5A5A5", "#E3E3E3", "#50AACD", "#0082B4"])

# Pandas-Library Setup
pd.options.mode.chained_assignment = None

# WebApp Kopf erstellen
col01, col02, col03 = st.columns((1, 6, 1))
with col02:
    st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')

# Globale Schriftart definieren
st.markdown(
    """
        <style>
        @font-face {L
        font-family: 'voestalpine';
        font-style: normal;
        font-weight: 400;
        }
        html, body, [class*="css"]  {
        font-family: 'voestalpine';
        font-size: 20px;
        }
        </style>
        """,
    unsafe_allow_html=True,
)

# Horizontales Optionsmenü erstellen
selected_hor = option_menu(
    menu_title=None,
    options=("Overview", "Effizienz", "ABC", "Portfolio", "Stückliste"),
    icons=("bezier", "bar-chart-fill", "calculator-fill",
           "grid-3x3-gap-fill", "card-list"),
    default_index=0,
    menu_icon="cast",
    orientation="horizontal",
    styles={
        "container": {"padding": "0!important", "background-color": "#262730"},
        "icon": {"color": "#FAFAFA"},
        "nav-link": {"font_family": "voestalpine", "text-align": "center",
                     "--hover-color": "#A5A5A5"},
        "nav-link-selected": {"background-color": "#0082B4"},
    }
)

# Spalten definieren
col1, col2 = st.columns((3, 1))

# Option Overview
if selected_hor == "Overview":

    # Überschrift erstellen
    with col1:
        st.markdown("""
            <style>
            .big-font {
            font-size:50px !important;
            color: #0082B4
            }
            </style>
            """, unsafe_allow_html=True)
        st.markdown('<p class="big-font">Overview</p>', unsafe_allow_html=True)

    # Railway Systems Logo einfügen
    with col2:
        st.image('https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

    # Datei-Uploader einfügen
    with st.expander("Upload"):
        iFiles = st.file_uploader(
            "", accept_multiple_files=True, type=["xlsm"])

    # Dateien nach Benennung sortieren
    iFiles.sort(key=lambda x: x.name.split("_")[2])

    # Produktnamen filtern
    iNames = []
    for i in iFiles:
        iNames.append(i.name.split("_")[0]+i.name.split("_")[2])

    # Produktnamen in Liste umwandeln
    iNames = list(dict.fromkeys(iNames))

    # Produktauswahl
    iFile = st.selectbox("Analyse-File Auswahl", iNames)

    # Auswahl zuweisen
    filesInput = []
    for i in iFiles:
        if i.name.split("_")[0]+i.name.split("_")[2] == iFile:
            filesInput.append(i)

    # Dashboard bei Datei-Upload
    if filesInput is not None:

        try:
            # Sidebar definieren
            with st.sidebar:
                # VoestAlpine Logo einfügen
                st.image(
                    'https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')
                # Sidebar Optionsmenü erstellen
                selected_side = option_menu(
                    menu_title="Materialgruppe",
                    options=("Allgemein", "Fertigungsteile",
                             "Zukaufteile", "Normteile"),
                    icons=("caret-right-fill", "caret-right-fill",
                           "caret-right-fill", "caret-right-fill"),
                    default_index=0,
                    menu_icon="cast",
                    styles={
                        "container": {"padding": "0!important", "background-color": "#262730"},
                        "icon": {"color": "#FAFAFA"},
                        "nav-link": {"font_family": "voestalpine", "text-align": "left",
                                     "--hover-color": "#A5A5A5"},
                        "nav-link-selected": {"background-color": "#0082B4"},
                    }
                )

                # Auswahl zuordnen
                if selected_side == "Fertigungsteile":
                    sheet = "Komponenten_FERT"
                if selected_side == "Zukaufteile":
                    sheet = "Komponenten_HIBE"
                if selected_side == "Normteile":
                    sheet = "Komponenten_NORM"
                if selected_side == "Allgemein":
                    sheet = "Komponenten"

                # Overview Auswahl
                files = filesInput

            # Ausgewählte Dataframes kombinieren
            dfs = {}
            for file in files:
                dfs[f"{file.name}"] = pd.read_excel(file,
                                                    engine="openpyxl",
                                                    sheet_name=sheet,
                                                    na_filter=True)

            # Spalte hinzufügen für Skalierung der X-Achse
            fill = []
            for file in files:
                for i in range(len(dfs[file.name].index)):
                    fill.append(i/(len(dfs[file.name].index)))
                dfs[file.name]["Prozent"] = fill
                fill.clear()

            # Universal-Anteil-Liste
            unipL = []
            # Exklusiv-Anteil-Liste
            exklpL = []

            # Parameter ermitteln
            for file in files:
                # Anzahl der Einzelkomponenten
                ein = len(dfs[file.name].index)
                unicount = 0
                exklcount = 0

                # Universalkomponenten zählen
                for i in range(ein):
                    if dfs[file.name].at[i, "Effizienz"] == 1:
                        unicount = unicount+1

                # Universalanteil ermitteln und der Liste zuweisen
                unip = unicount/ein
                unipL.append(unip)

                # Exklusivkomponenten zählen
                for i in range(ein):
                    if dfs[file.name].at[i, "Abfrage"] == 1:
                        exklcount = exklcount+1

                # Exklusivanteil ermitteln und der Liste zuweisen
                exklp = exklcount/ein
                exklpL.append(exklp)

            # Mittelwerte ermitteln
            SumUni = sum(unipL)
            avUni = (SumUni/len(unipL))*100
            SumExkl = sum(exklpL)
            avExkl = (SumExkl/len(exklpL))*100

            # Fertigungsteile
            dfsPF = {}
            for file in files:
                dfsPF[f"{file.name}"] = pd.read_excel(file,
                                                      engine="openpyxl",
                                                      sheet_name="Komponenten_FERT",
                                                      na_filter=False)

            PF = []
            for file in files:
                PF.append(len(dfsPF[file.name].index))

            SumFERT = sum(PF)

            # Zukaufteile
            dfsPH = {}
            for file in files:
                dfsPH[f"{file.name}"] = pd.read_excel(file,
                                                      engine="openpyxl",
                                                      sheet_name="Komponenten_HIBE",
                                                      na_filter=False)

            PH = []
            for file in files:
                PH.append(len(dfsPH[file.name].index))

            SumHIBE = sum(PH)

            # Normteile
            dfsPN = {}
            for file in files:
                dfsPN[f"{file.name}"] = pd.read_excel(file,
                                                      engine="openpyxl",
                                                      sheet_name="Komponenten_NORM",
                                                      na_filter=False)

            PN = []
            for file in files:
                PN.append(len(dfsPN[file.name].index))

            SumNORM = sum(PN)

            # General
            dfsP = {}
            for file in files:
                dfsP[f"{file.name}"] = pd.read_excel(file,
                                                     engine="openpyxl",
                                                     sheet_name=sheet,
                                                     na_filter=False)

            # Gesamtsumme
            SumGEN = SumFERT+SumHIBE+SumNORM

            # Anteilermittlung
            AnFERT = SumFERT/SumGEN
            AnHIBE = SumHIBE/SumGEN
            AnNORM = SumNORM/SumGEN

            # Mittelwerte ermitteln
            MW = []
            for i in dfsP:
                MW.append(sum(dfsP[i]["Effizienz"])/len(dfsP[i].index))

            # Mittelwerte benennen
            MWN = []
            for file in files:
                MWN.append("von " + file.name.split("_")[1])

            # Komponenten ohne Materialart entfernen
            for file in files:
                dfs[file.name].dropna(subset=["Materialart"], inplace=True)

            # Dashgrid definieren
            figD = make_subplots(
                rows=1, cols=2,
                specs=[[{"colspan": 2}, None]]
            )

            # Liniendiagramm (Bauteileffizienz) einfügen
            for i in dfs:
                figD.add_trace(go.Scatter(x=dfs[i]["Prozent"],
                                          y=dfs[i]["Effizienz"],
                                          name="von " + i.split("_")[1],
                                          mode='lines',
                                          line_color=next(palette)),
                               row=1, col=1)

            figD.update_xaxes(showgrid=False,
                              showticklabels=False,
                              row=1, col=1)
            figD.update_yaxes(showgrid=False,
                              title=None,
                              tickfont_size=20,
                              titlefont_size=30,
                              titlefont_family="voestalpine",
                              tickformat=',.0%',
                              row=1, col=1)
            figD.update_traces(showlegend=True,
                               hoverinfo="name+y",
                               hoverlabel_namelength=-1,
                               row=1, col=1)

            # Dashlayout anpassen
            figD.update_layout(plot_bgcolor="rgba(256,256,256,0)",
                               title={"text": "Bauteileffizienz",
                                      'y': 0.9,
                                      'x': 0.5,
                                      'xanchor': 'center',
                                      'yanchor': 'top'},
                               title_font_family="voestalpine",
                               title_font_color="#0082B4",
                               title_font_size=50,
                               legend_font=dict(family="voestalpine", size=18),
                               legend=dict(
                                   orientation="h",
                                   yanchor="top",
                                   y=0.05,
                                   xanchor="left",
                                   x=0.01,
                                   bgcolor="rgba(256,256,256,0)"),
                               height=400)

            # Dash anzeigen
            st.plotly_chart(figD, use_container_width=True)

            # In Prozentwerte umwandeln
            MWs = []
            for i in MW:
                MWs.append(float("{:.1f}".format(i*100)))

            # Durchschnittliche Effizienz
            optionB = {
                "tooltip": {
                    "show": True,
                    "formatter": """{b}: {c} %""",
                    "borderColor": "#0082B4",
                    "borderWidth": 3,
                    "backgroundColor": "#0E1117",
                    "textStyle": {
                        "color": "#FFF",
                        "fontFamily": "voestalpine",
                        "fontSize": 20
                    },
                    "trigger": "axis",
                    "axisPointer": {
                        "type": "shadow",
                        "shadowStyle": {
                            "color": "#0082B4",
                            "opacity": 0.5
                        }
                    }
                },
                "title": {
                    "text": 'Durchschnittliche Effizienz',
                    "left": 'center',
                    "top": 0,
                    "textStyle": {
                        "fontFamily": "voestalpine",
                        "fontWeight": "normal",
                        "fontSize": 50,
                        "color": '#0082B4'
                    }
                },
                "xAxis": {
                    "axisLabel": {"color": "#FFF",
                                  "fontFamily": "voestalpine",
                                  "fontWeight": "normal",
                                  "fontSize": 20},
                    "type": 'category',
                    "data": MWN
                },
                "yAxis": {
                    "type": 'value',
                    "show": False
                },
                "series": [
                    {
                        "data": MWs,
                        "label": {
                            "show": True,
                            "color": "white",
                            "position": "top",
                            "fontFamily": "voestalpine",
                            "fontSize": 25,
                            # "rotate":90,
                            "padding": [0, 0, -40, 0],
                            "align":"center",
                            "formatter": """{c}%"""
                        },
                        "type": 'bar',
                        "color":"#0082B4"
                    },
                ]
            }

            st_echarts(optionB, height=300, width="100%")

            # Sidebar modifizieren
            with st.sidebar:
                st.write(" ")

                # Anteile ermitteln
                AnFERT = float("{:.1f}".format(AnFERT*100))
                AnHIBE = float("{:.1f}".format(AnHIBE*100))
                AnNORM = float("{:.1f}".format(AnNORM*100))

                # Daten zuweisen
                labels = ['FERT', 'HIBE', 'NORM']
                values = [AnFERT, AnHIBE, AnNORM]

                # Materialart-Aufteilung
                optionP = {
                    "backgroundColor": 'rgba(0, 0, 0, 0)',
                    "series": [
                        {
                            "name": 'Materialart',
                            "type": 'pie',
                            "radius": '70%',
                            "center": ['50%', '50%'],
                            "data": [
                                {"value": AnFERT, "name": 'FERT'},
                                {"value": AnHIBE, "name": 'HIBE'},
                                {"value": AnNORM, "name": 'NORM'},
                            ],
                            "label": {
                                "color": '#0082B4',
                                "formatter": '{b}\n{c}%',
                                "fontSize": 20,
                            },
                            "labelLine": {
                                "lineStyle": {
                                    "color": 'rgba(255, 255, 255, 0.3)'
                                },
                                "smooth": 0.2,
                                "length": 10,
                                "length2": 20
                            },
                            "color": px.colors.qualitative.VoestBlue,
                            "animationType": 'scale',
                            "animationEasing": 'elasticOut',
                            "animationDelay": "function (idx) {return Math.random() * 200;}"
                        }
                    ]
                }

                st_echarts(optionP, height=175, width="100%")

            # In Prozentwerte umwandeln
            avUni = float("{:.1f}".format(avUni))
            avExkl = float("{:.1f}".format(avExkl))

            # Spalten definieren
            colG1, colG2 = st.columns((1, 1))
            with colG1:
                # Universalanteil
                optionG1 = {
                    "series": [
                        {
                            "type": 'gauge',
                            "startAngle": 180,
                            "endAngle": 0,
                            "min": 0,
                            "max": 100,
                            "splitNumber": 5,
                            "itemStyle": {
                                "color": '#0082B4',
                                "shadowColor": 'rgba(0,138,255,0.45)',
                                "shadowBlur": 10,
                                "shadowOffsetX": 2,
                                "shadowOffsetY": 2
                            },
                            "progress": {
                                "show": True,
                                "roundCap": True,
                                "width": 18
                            },
                            "pointer": {
                                "icon": 'path://M2090.36389,615.30999 L2090.36389,615.30999 C2091.48372,615.30999 2092.40383,616.194028 2092.44859,617.312956 L2096.90698,728.755929 C2097.05155,732.369577 2094.2393,735.416212 2090.62566,735.56078 C2090.53845,735.564269 2090.45117,735.566014 2090.36389,735.566014 L2090.36389,735.566014 C2086.74736,735.566014 2083.81557,732.63423 2083.81557,729.017692 C2083.81557,728.930412 2083.81732,728.84314 2083.82081,728.755929 L2088.2792,617.312956 C2088.32396,616.194028 2089.24407,615.30999 2090.36389,615.30999 Z',
                                "length": '75%',
                                "width": 16,
                                "offsetCenter": [0, '5%']
                            },
                            "axisLine": {
                                "roundCap": True,
                                "lineStyle": {
                                    "width": 18,
                                }
                            },
                            "axisTick": {
                                "splitNumber": 2,
                                "lineStyle": {
                                    "width": 2,
                                    "color": '#999'
                                }
                            },
                            "splitLine": {
                                "length": 12,
                                "lineStyle": {
                                    "width": 3,
                                    "color": '#999'
                                }
                            },
                            "axisLabel": {
                                "distance": 30,
                                "color": '#999',
                                "fontSize": 20
                            },
                            "data": [
                                {
                                    "value": avUni,
                                    "name": "Universalanteil"
                                }
                            ],
                            "detail": {
                                "valueAnimation": True,
                                "formatter": '{value}%',
                                "color": "#FFF",
                                "fontFamily": "voestalpine",
                                "fontSize": 30,
                                "fontWeight": "normal",
                            },
                            "title": {
                                "show": False,
                                "textStyle": {
                                    "fontFamily": "voestalpine",
                                    "fontWeight": "normal",
                                    "fontSize": 30,
                                    "color": '#0082B4'
                                }
                            }
                        }
                    ]
                }

                st_echarts(optionG1, height=400, width="100%")

            with colG2:
                # Exklusivanteil
                optionG2 = {
                    "series": [
                        {
                            "type": 'gauge',
                            "startAngle": 180,
                            "endAngle": 0,
                            "min": 0,
                            "max": 100,
                            "splitNumber": 5,
                            "itemStyle": {
                                "color": '#0082B4',
                                "shadowColor": 'rgba(0,138,255,0.45)',
                                "shadowBlur": 10,
                                "shadowOffsetX": 2,
                                "shadowOffsetY": 2
                            },
                            "progress": {
                                "show": True,
                                "roundCap": True,
                                "width": 18
                            },
                            "pointer": {
                                "icon": 'path://M2090.36389,615.30999 L2090.36389,615.30999 C2091.48372,615.30999 2092.40383,616.194028 2092.44859,617.312956 L2096.90698,728.755929 C2097.05155,732.369577 2094.2393,735.416212 2090.62566,735.56078 C2090.53845,735.564269 2090.45117,735.566014 2090.36389,735.566014 L2090.36389,735.566014 C2086.74736,735.566014 2083.81557,732.63423 2083.81557,729.017692 C2083.81557,728.930412 2083.81732,728.84314 2083.82081,728.755929 L2088.2792,617.312956 C2088.32396,616.194028 2089.24407,615.30999 2090.36389,615.30999 Z',
                                "length": '75%',
                                "width": 16,
                                "offsetCenter": [0, '5%']
                            },
                            "axisLine": {
                                "roundCap": True,
                                "lineStyle": {
                                    "width": 18
                                }
                            },
                            "axisTick": {
                                "splitNumber": 2,
                                "lineStyle": {
                                    "width": 2,
                                    "color": '#999'
                                }
                            },
                            "splitLine": {
                                "length": 12,
                                "lineStyle": {
                                    "width": 3,
                                    "color": '#999'
                                }
                            },
                            "axisLabel": {
                                "distance": 30,
                                "color": '#999',
                                "fontSize": 20
                            },
                            "data": [
                                {
                                    "value": avExkl,
                                    "name": "Exklusivanteil"
                                }
                            ],
                            "detail": {
                                "valueAnimation": True,
                                "formatter": '{value}%',
                                "color": "#FFF",
                                "fontFamily": "voestalpine",
                                "fontSize": 30,
                                "fontWeight": "normal",
                            },
                            "title": {
                                "show": False,
                                "textStyle": {
                                    "fontFamily": "voestalpine",
                                    "fontWeight": "normal",
                                    "fontSize": 30,
                                    "color": '#0082B4'
                                }
                            }
                        }
                    ]
                }

                st_echarts(optionG2, height=400, width="100%")

            # Parameter ermitteln und initialisieren
            dfd = {}
            for file in files:
                dfd[f"{file.name}"] = pd.read_excel(file,
                                                    engine="openpyxl",
                                                    sheet_name="Eingabe",
                                                    na_filter=False,
                                                    usecols="A")

            # Variantenwerte zuweisen
            var = []
            for i in dfd:
                var.append(len(dfd[i].index))

            # Deltawerte berechnen
            d0 = var[3]
            d1 = var[0]-var[1]
            d2 = var[1]-var[2]
            d3 = var[2]-var[3]

            # Formtierung zuweisen
            var1 = f"""<style>p.a {{font:50px voestalpine;color: #FFF;text-align: center}}</style><p class="a">{d1}</p>"""
            var2 = f"""<style>p.a {{font:50px voestalpine;color: #FFF;text-align: center}}</style><p class="a">{d1+d2}</p>"""
            var3 = f"""<style>p.a {{font:50px voestalpine;color: #FFF;text-align: center}}</style><p class="a">{d1+d2+d3}</p>"""
            var4 = f"""<style>p.a {{font:50px voestalpine;color: #FFF;text-align: center}}</style><p class="a">{d1+d2+d3+d0}</p>"""

            r1 = f"""<style>p.b {{font:30px voestalpine;color: #0082B4;text-align: center}}</style><p class="b">+{d1}</p>"""
            r2 = f"""<style>p.b {{font:30px voestalpine;color: #0082B4;text-align: center}}</style><p class="b">+{d2}</p>"""
            r3 = f"""<style>p.b {{font:30px voestalpine;color: #0082B4;text-align: center}}</style><p class="b">+{d3}</p>"""
            r4 = f"""<style>p.b {{font:30px voestalpine;color: #0082B4;text-align: center}}</style><p class="b">+{d0}</p>"""

            # Bereichüberschrift
            st.markdown(
                '<p style="text-align: center;color: #0082B4;font-size:50px">Variantenentwicklung</p>', unsafe_allow_html=True)

            # Spalten definieren
            colvo, colv1, colv2, colv3, colv4, colv5 = st.columns(
                (1, 2, 2, 2, 2, 1))

            # Werte ausgeben
            with colv1:
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">2018</p>', unsafe_allow_html=True)
                st.markdown(var1, unsafe_allow_html=True)
                st.markdown(r1, unsafe_allow_html=True)

            with colv2:
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">2019</p>', unsafe_allow_html=True)
                st.markdown(var2, unsafe_allow_html=True)
                st.markdown(r2, unsafe_allow_html=True)

            with colv3:
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">2020</p>', unsafe_allow_html=True)
                st.markdown(var3, unsafe_allow_html=True)
                st.markdown(r3, unsafe_allow_html=True)

            with colv4:
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">2021</p>', unsafe_allow_html=True)
                st.markdown(var4, unsafe_allow_html=True)
                st.markdown(r4, unsafe_allow_html=True)

            # werte sortieren
            var.sort()
            MWN.sort(reverse=True)

            # Spalten definieren
            colvf1, colvf2, colvf3 = st.columns((1, 50, 1))

            # Variantenentwicklung
            optionv = {
                "tooltip": {
                    "show": False,
                    "formatter": """{b}""",
                    "borderColor": "#FFF",
                    "borderWidth": 3,
                    "backgroundColor": "#0E1117",
                    "textStyle": {
                        "color": "#0082B4",
                        "fontFamily": "voestalpine",
                        "fontSize": 20
                    },
                    "trigger": "axis",
                    "axisPointer": {
                        "type": "shadow",
                        "shadowStyle": {
                            "color": "#0082B4",
                            "opacity": 0.5
                        }
                    }
                },
                "xAxis": {
                    "type": 'category',
                    "axisLabel": {"show": False},
                    "data": MWN
                },
                "yAxis": {
                    "type": 'value',
                    "show": False
                },
                "series": [
                    {
                        "data": [0, d1, d1+d2, d1+d2+d3],
                        "xAxisIndex": 0,
                        "yAxisIndex": 0,
                        "type": 'bar',
                        "stack": "total",
                        "color":"#FFF",
                        "itemStyle":{
                            "decal": {
                                "symbol": "rect",
                                "symbolSize": 0.95,
                                "color": "#0082B4",
                                "dashArrayX": [1, 0],
                                "dashArrayY":[2, 4],
                                "rotation": 0.785
                            }
                        }
                    },
                    {
                        "data": [d1, d2, d3, d0],
                        "xAxisIndex": 0,
                        "yAxisIndex": 0,
                        "type": 'bar',
                        "stack": "total",
                        "color":"#0082B4"
                    }
                ]
            }

            st_echarts(optionv, height=300, width="100%")

            # Bereichüberschrift
            st.write(" ")
            st.markdown(
                '<p style="text-align: center;color: #0082B4;font-size:50px">Portfolio</p>', unsafe_allow_html=True)

            # Spalten definieren
            col111, col112, col113, col114 = st.columns((6, 5, 5, 5))
            col11, col12 = st.columns((2, 5))

            # Spaltenbeschriftung
            with col112:
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:35px">Klasse A</p>', unsafe_allow_html=True)

            with col113:
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:35px">Klasse B</p>', unsafe_allow_html=True)

            with col114:
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:35px">Klasse C</p>', unsafe_allow_html=True)

            # Zeilenbeschriftung
            with col11:
                st. write("")
                st. write("")
                st. write("")
                st. write("")
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">Top Effizienz</p>', unsafe_allow_html=True)
                st.markdown(
                    '<p style="text-align: center;color: #FFF;font-size:26px">66 - 100%</p>', unsafe_allow_html=True)
                st. write("")
                st. write("")
                st. write("")
                st. write("")
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">Medium Effizienz</p>', unsafe_allow_html=True)
                st.markdown(
                    '<p style="text-align: center;color: #FFF;font-size:26px">33 - 66%</p>', unsafe_allow_html=True)
                st. write("")
                st. write("")
                st. write("")
                st. write("")
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">Low Effizienz</p>', unsafe_allow_html=True)
                st.markdown(
                    '<p style="text-align: center;color: #FFF;font-size:26px">0 - 33%</p>', unsafe_allow_html=True)

            # Werte zuweisen
            dfP = {}
            for i in files:
                dfP[f"{i.name}"] = pd.read_excel(i,
                                                 engine="openpyxl",
                                                 sheet_name="Portfolio_Auswertung",
                                                 na_filter=True)

            # Werte in Bereiche aufteilen
            Header0 = []
            for i in files:
                Header0.append(i.name.split("_")[1])

            Header = []
            for i in files:
                Header.append(i.name.split("_")[1])

            Low_A = []
            for i in files:
                Low_A.append(float("{:.1f}".format(
                    dfP[i.name].at[0, "Anteil"]*100.0)))

            Medium_A = []
            for i in files:
                Medium_A.append(float("{:.1f}".format(
                    dfP[i.name].at[1, "Anteil"]*100.0)))

            Top_A = []
            for i in files:
                Top_A.append(float("{:.1f}".format(
                    dfP[i.name].at[2, "Anteil"]*100.0)))

            Low_B = []
            for i in files:
                Low_B.append(float("{:.1f}".format(
                    dfP[i.name].at[3, "Anteil"]*100.0)))

            Medium_B = []
            for i in files:
                Medium_B.append(float("{:.1f}".format(
                    dfP[i.name].at[4, "Anteil"]*100.0)))

            Top_B = []
            for i in files:
                Top_B.append(float("{:.1f}".format(
                    dfP[i.name].at[5, "Anteil"]*100.0)))

            Low_C = []
            for i in files:
                Low_C.append(float("{:.1f}".format(
                    dfP[i.name].at[6, "Anteil"]*100.0)))

            Medium_C = []
            for i in files:
                Medium_C.append(float("{:.1f}".format(
                    dfP[i.name].at[7, "Anteil"]*100.0)))

            Top_C = []
            for i in files:
                Top_C.append(float("{:.1f}".format(
                    dfP[i.name].at[8, "Anteil"]*100.0)))

            # Portfolio
            option = {
                "tooltip": {
                    "show": True,
                    "formatter": """von {b}: {c} %""",
                    "borderColor": "#0082B4",
                    "borderWidth": 3,
                    "backgroundColor": "#0E1117",
                    "textStyle": {
                        "color": "#FFF",
                        "fontFamily": "voestalpine",
                        "fontSize": 20
                    },
                    "trigger": "axis",
                    "axisPointer": {
                        "type": "shadow",
                        "shadowStyle": {
                            "color": "#0082B4",
                            "opacity": 0.5
                        }
                    }
                },
                "xAxis": [
                    {"type": 'category', "gridIndex": 0, "data": Header},
                    {"type": 'category', "gridIndex": 1, "data": Header},
                    {"type": 'category', "gridIndex": 2, "data": Header},
                    {"type": 'category', "gridIndex": 3, "data": Header},
                    {"type": 'category', "gridIndex": 4, "data": Header},
                    {"type": 'category', "gridIndex": 5, "data": Header},
                    {"type": 'category', "gridIndex": 6, "data": Header},
                    {"type": 'category', "gridIndex": 7, "data": Header},
                    {"type": 'category', "gridIndex": 8, "data": Header}
                ],
                "yAxis": [
                    {"gridIndex": 0, "min": 0, "max": 100, "show": False},
                    {"gridIndex": 1, "min": 0, "max": 100, "show": False},
                    {"gridIndex": 2, "min": 0, "max": 100, "show": False},
                    {"gridIndex": 3, "min": 0, "max": 100, "show": False},
                    {"gridIndex": 4, "min": 0, "max": 100, "show": False},
                    {"gridIndex": 5, "min": 0, "max": 100, "show": False},
                    {"gridIndex": 6, "min": 0, "max": 100, "show": False},
                    {"gridIndex": 7, "min": 0, "max": 100, "show": False},
                    {"gridIndex": 8, "min": 0, "max": 100, "show": False},
                ],
                "grid": [
                    {"top": '0%', "left": "0%", "width": "30%", "height": "30%"},
                    {"top": '0%', "left": "35%", "width": "30%", "height": "30%"},
                    {"top": '0%', "right": "0%", "width": "30%", "height": "30%"},
                    {"top": '33.5%', "left": "0%", "width": "30%", "height": "30%"},
                    {"top": '33.5%', "left": "35%",
                        "width": "30%", "height": "30%"},
                    {"top": '33.5%', "right": "0%",
                        "width": "30%", "height": "30%"},
                    {"bottom": '3%', "left": "0%", "width": "30%", "height": "30%"},
                    {"bottom": '3%', "left": "35%",
                        "width": "30%", "height": "30%"},
                    {"bottom": '3%', "right": "0%", "width": "30%", "height": "30%"}
                ],
                "series": [
                    {
                        "data": Top_A,
                        "label": {
                            "show": True,
                            "color": "white",
                            "position": "top",
                            "fontFamily": "voestalpine",
                            "fontSize": 20,
                            "rotate": 90,
                            "padding": [0, 0, -9, 5],
                            "align":"left",
                            "formatter": """{c} %"""
                        },
                        "xAxisIndex": 0,
                        "yAxisIndex": 0,
                        "type": 'bar',
                        "color":"#0eff00"
                    },
                    {
                        "data": Top_B,
                        "label": {
                            "show": True,
                            "color": "white",
                            "position": "top",
                            "fontFamily": "voestalpine",
                            "fontSize": 20,
                            "rotate": 90,
                            "padding": [0, 0, -9, 5],
                            "align":"left",
                            "formatter": """{c} %"""
                        },
                        "xAxisIndex": 1,
                        "yAxisIndex": 1,
                        "type": 'bar',
                        "color":"yellow"
                    },
                    {
                        "data": Top_C,
                        "label": {
                            "show": True,
                            "color": "white",
                            "position": "top",
                            "fontFamily": "voestalpine",
                            "fontSize": 20,
                            "rotate": 90,
                            "padding": [0, 0, -9, 5],
                            "align":"left",
                            "formatter": """{c} %"""
                        },
                        "xAxisIndex": 2,
                        "yAxisIndex": 2,
                        "type": 'bar',
                        "color":"orange"
                    },
                    {
                        "data": Medium_A,
                        "label": {
                            "show": True,
                            "color": "white",
                            "position": "top",
                            "fontFamily": "voestalpine",
                            "fontSize": 20,
                            "rotate": 90,
                            "padding": [0, 0, -9, 5],
                            "align":"left",
                            "formatter": """{c} %"""
                        },
                        "xAxisIndex": 3,
                        "yAxisIndex": 3,
                        "type": 'bar',
                        "color":"yellow"
                    },
                    {
                        "data": Medium_B,
                        "label": {
                            "show": True,
                            "color": "white",
                            "position": "top",
                            "fontFamily": "voestalpine",
                            "fontSize": 20,
                            "rotate": 90,
                            "padding": [0, 0, -9, 5],
                            "align":"left",
                            "formatter": """{c} %"""
                        },
                        "xAxisIndex": 4,
                        "yAxisIndex": 4,
                        "type": 'bar',
                        "color":"orange"
                    },
                    {
                        "data": Medium_C,
                        "label": {
                            "show": True,
                            "color": "white",
                            "position": "top",
                            "fontFamily": "voestalpine",
                            "fontSize": 20,
                            "rotate": 90,
                            "padding": [0, 0, -9, 5],
                            "align":"left",
                            "formatter": """{c} %"""
                        },
                        "xAxisIndex": 5,
                        "yAxisIndex": 5,
                        "type": 'bar',
                        "color":"red"
                    },
                    {
                        "data": Low_A,
                        "label": {
                            "show": True,
                            "color": "white",
                            "position": "top",
                            "fontFamily": "voestalpine",
                            "fontSize": 20,
                            "rotate": 90,
                            "padding": [0, 0, -9, 5],
                            "align":"left",
                            "formatter": """{c} %"""
                        },
                        "xAxisIndex": 6,
                        "yAxisIndex": 6,
                        "type": 'bar',
                        "color":"orange"
                    },
                    {
                        "data": Low_B,
                        "label": {
                            "show": True,
                            "color": "white",
                            "position": "top",
                            "fontFamily": "voestalpine",
                            "fontSize": 20,
                            "rotate": 90,
                            "padding": [0, 0, -9, 5],
                            "align":"left",
                            "formatter": """{c} %"""
                        },
                        "xAxisIndex": 7,
                        "yAxisIndex": 7,
                        "type": 'bar',
                        "color":"red"
                    },
                    {
                        "data": Low_C,
                        "label": {
                            "show": True,
                            "color": "white",
                            "position": "top",
                            "fontFamily": "voestalpine",
                            "fontSize": 20,
                            "rotate": 90,
                            "padding": [0, 0, -9, 5],
                            "align":"left",
                            "formatter": """{c} %"""
                        },
                        "xAxisIndex": 8,
                        "yAxisIndex": 8,
                        "type": 'bar',
                        "color":"red"
                    }
                ]
            }

            with col12:
                st_echarts(option, height=620)

        # Bei fehlenden Input:
        except ZeroDivisionError:
            st.markdown(
                '<p style="text-align: center;color: #FFF;font-size:20px">Bitte uploaden Sie die Analyse-Dateien!</p>', unsafe_allow_html=True)

# Option ABC-Analyse
if selected_hor == "ABC":

    try:
        # Überschrift erstellen
        with col1:
            st.markdown("""
                <style>
                .big-font {
                font-size:50px !important;
                color: #0082B4
                }
                </style>
                """, unsafe_allow_html=True)
            st.markdown('<p class="big-font">ABC-Analyse</p>',
                        unsafe_allow_html=True)

        # Railway Systems Logo einfügen
        with col2:
            st.image(
                'https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

        # Datei-Uploader einfügen
        with st.expander("Upload"):
            iFiles = st.file_uploader(
                "", accept_multiple_files=True, type=["xlsm"])

        # Initialisierung
        file = None

        # Dateien nach Benennung sortieren
        iFiles.sort(key=lambda x: x.name.split("_")[2])

        # Produktnamen filtern
        iNames = []
        for i in iFiles:
            iNames.append(i.name.split("_")[0]+i.name.split("_")[2])

        # Produktnamen in Liste umwandeln
        iNames = list(dict.fromkeys(iNames))

        # Produktauswahl
        iFile = st.selectbox("Analyse-File Auswahl", iNames, key="ABCV")

        # Auswahl zuweisen
        filesT = []
        for i in iFiles:
            if i.name.split("_")[0]+i.name.split("_")[2] == iFile:
                filesT.append(i)

        # Jahrauswahl
        option1 = {
            "dataset": {},
            "tooltip": {
                "show": False,
            },
            "series": [
                {
                    "type": 'liquidFill',
                    "name": "2018",
                    "center": ["20%", "50%"],
                    "data":[1],
                    "itemStyle": {
                        "color": "white",
                        "shadowBlur": 0
                    },
                    "amplitude": 0,
                    "label": {
                        "formatter": "{a}",
                        "insideColor": "#0082B4",
                        "fontSize": "20"
                    },
                    "radius": "90%",
                    "waveAnimation": 0,
                    "shape": 'roundRect',
                    "backgroundStyle": {
                        "color": "grey"
                    },
                    "outline": {
                        "borderDistance": 5,
                        "itemStyle": {
                            "borderWidth": 5,
                            "borderColor": '#0082B4',
                            "shadowBlur": 0,
                            "shadowColor": 'green'
                        }}},
                {
                    "type": 'liquidFill',
                    "name": "2019",
                    "center": ["40%", "50%"],
                    "data":[1],
                    "itemStyle": {
                        "color": "white",
                        "shadowBlur": 0
                    },
                    "amplitude": 0,
                    "label": {
                        "formatter": "{a}",
                        "insideColor": "#0082B4",
                        "fontSize": "20"
                    },
                    "radius": "90%",
                    "waveAnimation": 0,
                    "shape": 'roundRect',
                    "backgroundStyle": {
                        "color": "grey"
                    },
                    "outline": {
                        "borderDistance": 5,
                        "itemStyle": {
                            "borderWidth": 5,
                            "borderColor": '#0082B4',
                            "shadowBlur": 0,
                            "shadowColor": 'green'
                        }}},
                {
                    "type": 'liquidFill',
                    "name": "2020",
                    "center": ["60%", "50%"],
                    "data":[1],
                    "itemStyle": {
                        "color": "white",
                        "shadowBlur": 0
                    },
                    "amplitude": 0,
                    "label": {
                        "formatter": "{a}",
                        "insideColor": "#0082B4",
                        "fontSize": "20"
                    },
                    "radius": "90%",
                    "waveAnimation": 0,
                    "shape": 'roundRect',
                    "backgroundStyle": {
                        "color": "grey"
                    },
                    "outline": {
                        "borderDistance": 5,
                        "itemStyle": {
                            "borderWidth": 5,
                            "borderColor": '#0082B4',
                            "shadowBlur": 0,
                            "shadowColor": 'green'
                        }}},
                {
                    "type": 'liquidFill',
                    "name": "2021",
                    "center": ["80%", "50%"],
                    "data":[1],
                    "itemStyle": {
                        "color": "white",
                        "shadowBlur": 0
                    },
                    "amplitude": 0,
                    "label": {
                        "formatter": "{a}",
                        "insideColor": "#0082B4",
                        "fontSize": "20"
                    },
                    "radius": "90%",
                    "waveAnimation": 0,
                    "shape": 'roundRect',
                    "backgroundStyle": {
                        "color": "grey"
                    },
                    "outline": {
                        "borderDistance": 5,
                        "itemStyle": {
                            "borderWidth": 5,
                            "borderColor": '#0082B4',
                            "shadowBlur": 0,
                            "shadowColor": 'green'
                        }}},
            ]
        }

        # Spalten definieren
        colv1, colv2, colv3 = st.columns((3, 2, 3))

        # Vergleichmodus
        with colv2:
            vergleich = st.checkbox("Vergleichmodus")

        # Click-Events
        events1 = {"click": "function(params) { return params.seriesName }"}

        # Click-Output
        output1 = st_echarts(option1, height=100, events=events1, key="1")

        # Jahrauswahl zuweisen
        for i in filesT:
            if i.name.split("_")[1] == output1:
                file = i

        # Jahrauswahl formatieren
        Jahr1 = f"""<style>p.a {{font:50px voestalpine;color: #0082B4;text-align: center}}</style><p class="a">von {output1}</p>"""

        # Jahrauswahl darstellen
        if output1 != None:
            st.markdown(Jahr1, unsafe_allow_html=True)

        # Bei Datei-Upload
        if file != None:
            dfr = pd.read_excel(file,
                                engine="openpyxl",
                                sheet_name="ABC",
                                na_filter=False,
                                usecols="A:D")

            dfv = pd.read_excel(file,
                                engine="openpyxl",
                                sheet_name="ABC",
                                na_filter=False,
                                usecols="G")

            dfg = pd.concat([dfr, dfv], axis=1, join="inner")

            # Spalten formatieren
            dfg['Wertanteil'] = dfr['Wertanteil'].map('{:,.1f}%'.format)
            dfg['Summe'] = dfr['Summe'].map('{:,.1f}%'.format)

            # ABC-Dataframe öffnen
            df = pd.read_excel(file,
                               engine="openpyxl",
                               sheet_name="ABC",
                               na_filter=False,
                               usecols="E:F",
                               header=None)

            # Spalte hinzufügen
            fill = []
            for i in range(len(dfr.index)+1):
                fill.append("-")

            df[6] = fill

            # Ersten drei Zeilen überschreiben
            df.at[0, 6] = "Klasse A"
            df.at[1, 6] = "Klasse B"
            df.at[2, 6] = "Klasse C"

            # Balkendiagramm (Mengenanteil) erstellen
            fig_bar1 = px.bar(df, y=4, color=6, text=4,
                              barmode='stack', orientation="v",
                              labels={"count": "Mengenanteil", "4": ""},
                              color_discrete_sequence=px.colors.qualitative.VoestBlue)

            # Balkendiagramm (Wertanteil) erstellen
            fig_bar2 = px.bar(df, y=5, color=6, text=5,
                              barmode='stack', orientation="v",
                              labels={"count": "Wertanteil", "5": ""},
                              color_discrete_sequence=px.colors.qualitative.VoestBlue)

            # Dashlayout 1 anpassen
            fig_bar1.update_layout(
                plot_bgcolor="rgba(256,256,256,0)",
                title_font_family="voestalpine",
                font_family="voestalpine",
                showlegend=False,
                xaxis=(dict(showgrid=False,
                            showticklabels=False,
                            titlefont_size=30)),
                yaxis=(dict(showgrid=False,
                            tickfont_size=15)),
                hoverlabel=dict(
                    bgcolor="#0082B4",
                    font_size=16,
                    font_family="voestalpine"
                )
            )

            # Dashlayout 2 anpassen
            fig_bar2.update_layout(
                plot_bgcolor="rgba(256,256,256,0)",
                title_font_family="voestalpine",
                font_family="voestalpine",
                showlegend=False,
                xaxis=(dict(showgrid=False,
                            showticklabels=False,
                            titlefont_size=30)),
                yaxis=(dict(showgrid=False,
                            tickfont_size=15)),
                hoverlabel=dict(
                    bgcolor="#0082B4",
                    font_size=16,
                    font_family="voestalpine"
                )
            )

            # Dashtraces 1 anpassen
            fig_bar1.update_traces(marker=dict(line=dict(width=0)),
                                   hovertemplate='%{y:.2f}%' +
                                   "<br>Mengenanteil",
                                   texttemplate='%{text:.2f}%',
                                   textposition="none",
                                   textfont=dict(
                family="voestalpine",
                size=25,))

            # Dashtraces 2 anpassen
            fig_bar2.update_traces(marker=dict(line=dict(width=0)),
                                   hovertemplate='%{y:.2f}%'+"<br>Wertanteil",
                                   texttemplate='%{text:.2f}%',
                                   textposition="none",
                                   textfont=dict(
                family="voestalpine",
                size=25,))

            # Spalten definieren
            col11, col12, col13 = st.columns((4, 2, 4))

            # Werte zuweisen
            testc1 = df.at[2, 4]
            testc2 = df.at[2, 5]
            testb1 = df.at[1, 4]
            testb2 = df.at[1, 5]
            testa1 = df.at[0, 4]
            testa2 = df.at[0, 5]

            # Zellen formatieren
            df[4].iloc[:3] = df[4].iloc[:3].map('{:,.1f}%'.format)
            df[5].iloc[:3] = df[5].iloc[:3].map('{:,.1f}%'.format)

            # Balkendiagramm (Mengenanteil) anzeigen
            with col11:
                st.plotly_chart(fig_bar1, use_container_width=True)

            # Werte anzeigen
            with col12:
                st.write("")
                st.write("")
                st.write("")
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">Klasse C</p>', unsafe_allow_html=True)
                st.write(df.at[2, 4]+" :arrow_right: "+df.at[2, 5])
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">Klasse B</p>', unsafe_allow_html=True)
                st.write(df.at[1, 4]+" :arrow_right: "+df.at[1, 5])
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">Klasse A</p>', unsafe_allow_html=True)
                st.write(df.at[0, 4]+" :arrow_right: "+df.at[0, 5])

            # Balkendiagramm (Wertanteil) anzeigen
            with col13:
                st.plotly_chart(fig_bar2, use_container_width=True)

            # Vergleichmodus
            if vergleich is True:
                # Jahrauswahl
                option2 = {
                    "dataset": {},
                    "tooltip": {
                        "show": False,
                    },
                    "series": [
                        {
                            "type": 'liquidFill',
                            "name": "2018",
                            "center": ["20%", "50%"],
                            "data":[1],
                            "itemStyle": {
                                "color": "white",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "formatter": "{a}",
                                "insideColor": "#0082B4",
                                "fontSize": "20"
                            },
                            "radius": "90%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "grey"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": '#0082B4',
                                    "shadowBlur": 0,
                                    "shadowColor": 'green'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "2019",
                            "center": ["40%", "50%"],
                            "data":[1],
                            "itemStyle": {
                                "color": "white",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "formatter": "{a}",
                                "insideColor": "#0082B4",
                                "fontSize": "20"
                            },
                            "radius": "90%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "grey"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": '#0082B4',
                                    "shadowBlur": 0,
                                    "shadowColor": 'green'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "2020",
                            "center": ["60%", "50%"],
                            "data":[1],
                            "itemStyle": {
                                "color": "white",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "formatter": "{a}",
                                "insideColor": "#0082B4",
                                "fontSize": "20"
                            },
                            "radius": "90%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "grey"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": '#0082B4',
                                    "shadowBlur": 0,
                                    "shadowColor": 'green'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "2021",
                            "center": ["80%", "50%"],
                            "data":[1],
                            "itemStyle": {
                                "color": "white",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "formatter": "{a}",
                                "insideColor": "#0082B4",
                                "fontSize": "20"
                            },
                            "radius": "90%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "grey"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": '#0082B4',
                                    "shadowBlur": 0,
                                    "shadowColor": 'green'
                                }}},
                    ]
                }

                # Click-Events
                events2 = {
                    "click": "function(params) { return params.seriesName }"}

                # Click-Output
                output2 = st_echarts(option2, height=100,
                                     events=events2, key="2")

                # Jahrauswahl zuweisen
                for i in filesT:
                    if i.name.split("_")[1] == output2:
                        file = i

                # Jahrauswahl formatieren
                Jahr2 = f"""<style>p.a {{font:50px voestalpine;color: #0082B4;text-align: center}}</style><p class="a">von {output2}</p>"""

                # Jahrauswahl darstellen
                if output2 is not None:
                    st.markdown(Jahr2, unsafe_allow_html=True)

                # Bei Dateiauswahl
                if file is not None:
                    dfr2 = pd.read_excel(file,
                                         engine="openpyxl",
                                         sheet_name="ABC",
                                         na_filter=False,
                                         usecols="A:D")

                    # Spalten formatieren
                    dfr2['Wertanteil'] = dfr2['Wertanteil'].map(
                        '{:,.1f}%'.format)
                    dfr2['Summe'] = dfr2['Summe'].map('{:,.1f}%'.format)

                    # ABC-Dataframe öffnen
                    df2 = pd.read_excel(file,
                                        engine="openpyxl",
                                        sheet_name="ABC",
                                        na_filter=False,
                                        usecols="E:F",
                                        header=None)

                    # Spalte hinzufügen
                    fill2 = []
                    for i in range(len(dfr2.index)+1):
                        fill2.append("-")

                    df2[6] = fill2

                    # Ersten drei Zeilen überschreiben
                    df2.at[0, 6] = "Klasse A"
                    df2.at[1, 6] = "Klasse B"
                    df2.at[2, 6] = "Klasse C"

                    # Balkendiagramm (Mengenanteil) erstellen
                    fig_bar12 = px.bar(df2, y=4, color=6, text=4,
                                       barmode='stack', orientation="v",
                                       labels={
                                           "count": "Mengenanteil", "4": ""},
                                       color_discrete_sequence=px.colors.qualitative.VoestBlue)

                    # Balkendiagramm (Wertanteil) erstellen
                    fig_bar22 = px.bar(df2, y=5, color=6, text=5,
                                       barmode='stack', orientation="v",
                                       labels={"count": "Wertanteil", "5": ""},
                                       color_discrete_sequence=px.colors.qualitative.VoestBlue)

                    # Dashlayout 1 anpassen
                    fig_bar12.update_layout(
                        plot_bgcolor="rgba(256,256,256,0)",
                        title_font_family="voestalpine",
                        font_family="voestalpine",
                        showlegend=False,
                        xaxis=(dict(showgrid=False,
                                    showticklabels=False,
                                    titlefont_size=30)),
                        yaxis=(dict(showgrid=False,
                                    tickfont_size=15)),
                        hoverlabel=dict(
                            bgcolor="#0082B4",
                            font_size=16,
                            font_family="voestalpine"
                        )
                    )

                    # Dashlayout 2 anpassen
                    fig_bar22.update_layout(
                        plot_bgcolor="rgba(256,256,256,0)",
                        title_font_family="voestalpine",
                        font_family="voestalpine",
                        showlegend=False,
                        xaxis=(dict(showgrid=False,
                                    showticklabels=False,
                                    titlefont_size=30)),
                        yaxis=(dict(showgrid=False,
                                    tickfont_size=15)),
                        hoverlabel=dict(
                            bgcolor="#0082B4",
                            font_size=16,
                            font_family="voestalpine"
                        )
                    )

                    # Dashtraces 1 anpassen
                    fig_bar12.update_traces(marker=dict(line=dict(width=0)),
                                            hovertemplate='%{y:.2f}%' +
                                            "<br>Mengenanteil",
                                            texttemplate='%{text:.2f}%',
                                            textposition="none",
                                            textfont=dict(
                        family="voestalpine",
                        size=25,))

                    # Dashtraces 2 anpassen
                    fig_bar22.update_traces(marker=dict(line=dict(width=0)),
                                            hovertemplate='%{y:.2f}%' +
                                            "<br>Wertanteil",
                                            texttemplate='%{text:.2f}%',
                                            textposition="none",
                                            textfont=dict(
                        family="voestalpine",
                        size=25,))

                    cm = "{:.1f}%".format(df2.at[2, 4]-testc1)
                    cw = "{:.1f}%".format(df2.at[2, 5]-testc2)

                    bm = "{:.1f}%".format(df2.at[1, 4]-testb1)
                    bw = "{:.1f}%".format(df2.at[1, 5]-testb2)

                    am = "{:.1f}%".format(df2.at[0, 4]-testa1)
                    aw = "{:.1f}%".format(df2.at[0, 5]-testa2)

                    # Zellen formatieren
                    df2[4].iloc[:3] = df2[4].iloc[:3].map('{:,.1f}%'.format)
                    df2[5].iloc[:3] = df2[5].iloc[:3].map('{:,.1f}%'.format)

                    colm1, colm2, colm3, colm4, colm5 = st.columns(
                        (2, 2, 2, 2, 2))

                    with colm4:
                        st.markdown(
                            '<p style="text-align: left;color: #0082B4;font-size:30px">Klasse C</p>', unsafe_allow_html=True)
                        st.metric("Mengenanteil:", df2.at[2, 4], cm)
                        st.metric("Wertanteil:", df2.at[2, 5], cw)

                    with colm3:
                        st.markdown(
                            '<p style="text-align: left;color: #0082B4;font-size:30px">Klasse B</p>', unsafe_allow_html=True)
                        st.metric("Mengenanteil:", df2.at[1, 4], bm)
                        st.metric("Wertanteil:", df2.at[1, 5], bw)

                    with colm2:
                        st.markdown(
                            '<p style="text-align: left;color: #0082B4;font-size:30px">Klasse B</p>', unsafe_allow_html=True)
                        st.metric("Mengenanteil:", df2.at[0, 4], am)
                        st.metric("Wertanteil:", df2.at[0, 5], aw)

                    # Spalten definieren
                    col112, col122, col132 = st.columns((4, 2, 4))

                    # Balkendiagramm (Mengenanteil) anzeigen
                    with col112:
                        st.plotly_chart(fig_bar12, use_container_width=True)

                    # Werte anzeigen
                    with col122:
                        st.write("")
                        st.write("")
                        st.write("")
                        st.markdown(
                            '<p style="text-align: center;color: #0082B4;font-size:30px">Klasse C</p>', unsafe_allow_html=True)
                        st.write(df2.at[2, 4]+" :arrow_right: "+df2.at[2, 5])
                        st.markdown(
                            '<p style="text-align: center;color: #0082B4;font-size:30px">Klasse B</p>', unsafe_allow_html=True)
                        st.write(df2.at[1, 4]+" :arrow_right: "+df2.at[1, 5])
                        st.markdown(
                            '<p style="text-align: center;color: #0082B4;font-size:30px">Klasse A</p>', unsafe_allow_html=True)
                        st.write(df2.at[0, 4]+" :arrow_right: "+df2.at[0, 5])

                    # Balkendiagramm (Wertanteil) anzeigen
                    with col132:
                        st.plotly_chart(fig_bar22, use_container_width=True)

            # Dataframe bei Einzelmodus
            else:
                gb = GridOptionsBuilder.from_dataframe(dfg)

                # Dataframe-Einstellungen
                grid_options = {
                    "defaultColDef": {
                        "minWidth": 5,
                        "editable": False,
                        "filter": True,
                        "resizable": True,
                        "sortable": True
                    },
                    "suppressFieldDotNotation": True,
                    "columnDefs": [
                        {
                            "headerName": "Materialnr.",
                            "field": "Materialnr.",
                            "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "18px"},
                            "width": 180,
                            "type": []
                        },
                        {
                            "headerName": "Variante",
                            "field": "Variante",
                            "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "18px"},
                            "width": 500,
                            "type": []
                        },
                        {
                            "headerName": "Wertanteil",
                            "field": "Wertanteil",
                            "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "18px"},
                            "width": 90,
                            "type": []
                        },
                        {
                            "headerName": "Summe",
                            "field": "Summe",
                            "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "18px"},
                            "width": 90,
                            "type": []
                        },
                        {
                            "headerName": "Klasse",
                            "field": "Klasse",
                            "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "18px"},
                            "width": 40,
                            "type": []
                        },
                    ]
                }

                # Dataframe anzeigen
                AgGrid(dfg, gridOptions=grid_options, theme="streamlit",
                       height=400, fit_columns_on_grid_load=True, editable=False)

    except ValueError:
        st.markdown('<p style="text-align: center;color: #FFF;font-size:20px">Die Input-Datei hat keine passende ABC-Aufschlüsselung!</p>', unsafe_allow_html=True)

# Option Effizienz-Analyse
if selected_hor == "Effizienz":

    # Überschrift einfügen
    with col1:
        st.markdown("""
            <style>
            .big-font {
            font-size:50px !important;
            color: #0082B4
            }
            </style>
            """, unsafe_allow_html=True)
        st.markdown('<p class="big-font">Effizienz-Analyse</p>',
                    unsafe_allow_html=True)

    # Railway Systems Logo einfügen
    with col2:
        st.image('https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

    # Datei-Uploader einfügen
    with st.expander("Upload"):
        iFiles = st.file_uploader(
            "", accept_multiple_files=True, type=["xlsm"])

    # Dateien nach Benennung sortieren
    iFiles.sort(key=lambda x: x.name.split("_")[2].split("-")[0])

    # Produktnamen filtern
    iNames = []
    for i in iFiles:
        iNames.append(i.name.split("_")[0]+i.name.split("_")[2])

    # Produktnamen in Liste umwandeln
    iNames = list(dict.fromkeys(iNames))

    # Produktauswahl
    iFile = st.selectbox("Analyse-File Auswahl", iNames, key="EFF-V")

    # Auswahl zuweisen
    files = []
    for i in iFiles:
        if i.name.split("_")[0]+i.name.split("_")[2] == iFile:
            files.append(i)

    # Initialisierung
    file = None

    # Jahrauswahl
    option1 = {
        "dataset": {},
        "tooltip": {
            "show": False,
        },
        "series": [
            {
                "type": 'liquidFill',
                "name": "2018",
                "center": ["20%", "50%"],
                "data":[1],
                "itemStyle": {
                    "color": "white",
                    "shadowBlur": 0
                },
                "amplitude": 0,
                "label": {
                    "formatter": "{a}",
                    "insideColor": "#0082B4",
                    "fontSize": "20"
                },
                "radius": "90%",
                "waveAnimation": 0,
                "shape": 'roundRect',
                "backgroundStyle": {
                    "color": "grey"
                },
                "outline": {
                    "borderDistance": 5,
                    "itemStyle": {
                        "borderWidth": 5,
                        "borderColor": '#0082B4',
                        "shadowBlur": 0,
                        "shadowColor": 'green'
                    }}},
            {
                "type": 'liquidFill',
                "name": "2019",
                "center": ["40%", "50%"],
                "data":[1],
                "itemStyle": {
                    "color": "white",
                    "shadowBlur": 0
                },
                "amplitude": 0,
                "label": {
                    "formatter": "{a}",
                    "insideColor": "#0082B4",
                    "fontSize": "20"
                },
                "radius": "90%",
                "waveAnimation": 0,
                "shape": 'roundRect',
                "backgroundStyle": {
                    "color": "grey"
                },
                "outline": {
                    "borderDistance": 5,
                    "itemStyle": {
                        "borderWidth": 5,
                        "borderColor": '#0082B4',
                        "shadowBlur": 0,
                        "shadowColor": 'green'
                    }}},
            {
                "type": 'liquidFill',
                "name": "2020",
                "center": ["60%", "50%"],
                "data":[1],
                "itemStyle": {
                    "color": "white",
                    "shadowBlur": 0
                },
                "amplitude": 0,
                "label": {
                    "formatter": "{a}",
                    "insideColor": "#0082B4",
                    "fontSize": "20"
                },
                "radius": "90%",
                "waveAnimation": 0,
                "shape": 'roundRect',
                "backgroundStyle": {
                    "color": "grey"
                },
                "outline": {
                    "borderDistance": 5,
                    "itemStyle": {
                        "borderWidth": 5,
                        "borderColor": '#0082B4',
                        "shadowBlur": 0,
                        "shadowColor": 'green'
                    }}},
            {
                "type": 'liquidFill',
                "name": "2021",
                "center": ["80%", "50%"],
                "data":[1],
                "itemStyle": {
                    "color": "white",
                    "shadowBlur": 0
                },
                "amplitude": 0,
                "label": {
                    "formatter": "{a}",
                    "insideColor": "#0082B4",
                    "fontSize": "20"
                },
                "radius": "90%",
                "waveAnimation": 0,
                "shape": 'roundRect',
                "backgroundStyle": {
                    "color": "grey"
                },
                "outline": {
                    "borderDistance": 5,
                    "itemStyle": {
                        "borderWidth": 5,
                        "borderColor": '#0082B4',
                        "shadowBlur": 0,
                        "shadowColor": 'green'
                    }}},
        ]
    }

    # Click-Events
    events1 = {"click": "function(params) { return params.seriesName }"}

    # Click-Output
    output1 = st_echarts(option1, height=100, events=events1, key="1")

    # Jahrauswahl formatieren
    Jahr1 = f"""<style>p.b{{font:50px voestalpine !important;color: #0082B4;text-align: center !important}}</style><p class="b">von {output1}</p>"""

    # Jahrauswahl darstellen
    if output1 is not None:
        st.markdown(Jahr1, unsafe_allow_html=True)

    # Jahrauswahl zuweisen
    for i in files:
        if i.name.split("_")[1] == output1:
            file = i

    # Bei Datei-Uplaod
    if file is not None:

        # Sidebar erstellen
        with st.sidebar:
            # VoestAlpine Logo einfügen
            st.image(
                'https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')
            # Sidebar Optionsmenü erstellen
            selected_side = option_menu(
                menu_title="Materialgruppe",
                options=("Allgemein", "Fertigungsteile",
                         "Zukaufteile", "Normteile"),
                icons=("caret-right-fill", "caret-right-fill",
                       "caret-right-fill", "caret-right-fill"),
                default_index=0,
                menu_icon="cast",
                styles={
                    "container": {"padding": "0!important", "background-color": "#262730"},
                    "icon": {"color": "#FAFAFA"},
                    "nav-link": {"font_family": "voestalpine", "text-align": "left", "--hover-color": "#A5A5A5"},
                    "nav-link-selected": {"background-color": "#0082B4"},
                }
            )

            # Auswahl zuordnen
            if selected_side == "Fertigungsteile":
                sheet = "Komponenten_FERT"
            if selected_side == "Zukaufteile":
                sheet = "Komponenten_HIBE"
            if selected_side == "Normteile":
                sheet = "Komponenten_NORM"
            if selected_side == "Allgemein":
                sheet = "Komponenten"

        # Eingabe-Dataframe öffnen
        dfd = pd.read_excel(file,
                            engine="openpyxl",
                            sheet_name="Eingabe",
                            na_filter=False,
                            usecols="A")

        # Ausgewählten Dataframe öffnen
        df = pd.read_excel(file,
                           engine="openpyxl",
                           sheet_name=sheet,
                           na_filter=True)

        # Parameter ermitteln und initialisieren
        var = len(dfd.index)
        ein = len(df.index)
        unicount = 0
        exklcount = 0

        # Universalkomponenten zählen
        for i in range(len(df.index)):
            if df.at[i, "Effizienz"] == 1:
                unicount = unicount+1

        # Exklusivkomponenten zählen
        for i in range(len(df.index)):
            if df.at[i, "Abfrage"] == 1:
                exklcount = exklcount+1

        uni = unicount
        exkl = exklcount

        # Anteile ermitteln
        unip = (unicount/ein)*100
        exklp = (exklcount/ein)*100

        # Spalte Effizienz formatieren
        df["Effizienz"] = df["Effizienz"]*100
        df['Effizienz'] = df['Effizienz'].map('{:,.1f}%'.format)

        # Spalte Mittelwert formatieren
        df["Mittelwert"] = df["Mittelwert"]*100
        df['Mittelwert'] = df['Mittelwert'].map('{:,.1f}%'.format)

        # Zeile einfügen
        df.iloc[0] = ['-', '-', '-', '-', '-', '-', '-',
                      '-', '-', '-', '-', '-', "0", "0", '-', '-']

        # Balkendiagramm (Bauteileffizienz) erstellen
        fig_area = px.bar(
            data_frame=df, y="Effizienz", x="Komponentennummer",
            labels=dict(Komponentennummer="Komponentenanzahl",
                        Effizienz="Bauteileffizienz"),
            height=500,
            range_y=(0, 100),
            base=None,
            hover_name="Objektsparte",
            # hover_data=["Abfrage"],
            custom_data=["Objektsparte", "Abfrage"],
            # color="Abfrage",
            # color_continuous_scale=px.colors.sequential.Oryel
            # color_discrete_sequence=px.colors.sequential.Plasma_r
            color_discrete_sequence=['#0082B4']*len(df)
        )

        # Dashtraces anpassen
        fig_area.update_traces(
            hovertemplate="<b>%{customdata[0]}</b><br>" +
            "Mat.Nr.: <b>%{x}</b><br>" +
            "Bauteileffizienz: <b>%{y}</b><br>" +
            "In <b>%{customdata[1]}</b> Varianten verbaut"
        )

        # Dashlayout anpassen
        fig_area.update_layout(
            plot_bgcolor="rgba(256,256,256,0)",
            title_font_family="voestalpine",
            font_family="voestalpine",
            xaxis=(dict(showgrid=False,
                        showticklabels=False,
                        tickfont_size=20,
                        titlefont_size=30)),
            yaxis=(dict(showgrid=False,
                        tickfont_size=20,
                        titlefont_size=30)),
        )

        # Y-Achse anpassen
        fig_area['layout']['yaxis'].update(autorange=True)

        # Dashtraces anpassen
        fig_area.update_traces(marker=dict(line=dict(width=0)))

        # Komponenten ohne Materialart entfernen
        df.dropna(subset=["Materialart"], inplace=True)

        # Dash anzeigen
        st.plotly_chart(fig_area, use_container_width=True)

        # Spalten definieren
        col11, col12, col13 = st.columns((2, 3, 6))

        # Schriftart definieren
        st.markdown("""
            <style>
            .font {
        font:30px voestalpine;
            }
            </style>
            """, unsafe_allow_html=True)

        # Textformatierung zuweisen
        unistr = f"""<style>p.a {{font:30px voestalpine !important;color: #0082B4;text-align: right !important}}</style><p class="a">{uni}</p>"""
        exklstr = f"""<style>p.a {{font:30px voestalpine !important;color: #0082B4;text-align: right !important}}</style><p class="a">{exkl}</p>"""
        varstr = f"""<style>p.a {{font:30px voestalpine !important;color: #0082B4;text-align: right !important}}</style><p class="a">{var}</p>"""
        einstr = f"""<style>p.a {{font:30px voestalpine !important;color: #0082B4;text-align: right !important}}</style><p class="a">{ein}</p>"""

        # Werte anzeigen
        with col11:
            st.markdown(unistr, unsafe_allow_html=True)
            st.markdown(exklstr, unsafe_allow_html=True)
            st.markdown(varstr, unsafe_allow_html=True)
            st.markdown(einstr, unsafe_allow_html=True)

        # Beschriftung anzeigen
        with col12:
            st.markdown('<p class="font">Universalkomp.</p>',
                        unsafe_allow_html=True)
            st.markdown('<p class="font">Exklusivkomp.</p>',
                        unsafe_allow_html=True)
            st.markdown('<p class="font">Varianten</p>',
                        unsafe_allow_html=True)
            st.markdown('<p class="font">Einzelkomp.</p>',
                        unsafe_allow_html=True)

        # Dataframe anpassen
        df = df.iloc[1:, :]

        gb = GridOptionsBuilder.from_dataframe(df)

        # Dataframe-Einstellungen
        grid_options = {
            "defaultColDef": {
                "minWidth": 5,
                "editable": False,
                "filter": True,
                "resizable": True,
                "sortable": True
            },
            "suppressFieldDotNotation": True,
            "columnDefs": [
                {
                    "headerName": "Komponentennummer",
                    "field": "Komponentennummer",
                    "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "15px"},
                    "width": 280,
                    "type": []
                },
                {
                    "headerName": "Effizienz",
                    "field": "Effizienz",
                    "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "15px"},
                    "width": 130,
                    "type": []
                },
                {
                    "headerName": "Materialart",
                    "field": "Materialart",
                    "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "15px"},
                    "width": 130,
                    "type": []
                },
                {
                    "headerName": "Objektkurztext",
                    "field": "Objektkurztext",
                    "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "15px"},
                    "width": 370,
                    "type": []
                },
            ]
        }

        # Dataframe anzeigen
        with col13:
            AgGrid(df, gridOptions=grid_options, theme="streamlit",
                   height=210, fit_columns_on_grid_load=True, editable=False)

        # Spalten definieren
        col111, col112, col113 = st.columns((8, 5, 8))

        # Zeigerdiagramm (Universalanteil) erstellen
        fig_uni = go.Figure(go.Indicator(
            domain={'x': [0, 1], 'y': [0, 1]},
            value=unip,
            mode="gauge+number",
            title={'text': "Universalanteil", 'font': {
                'size': 40, 'family': 'voestalpine'}},
            #delta = {'reference': 50},
            number={'suffix': "%"},
            gauge={
                "bar": {"color": "#0eff00"},
                'axis': {'range': [None, 100]}
            }))

        # Zeigerdiagramm (Exklusivanteil) erstellen
        fig_exkl = go.Figure(go.Indicator(
            domain={'x': [0, 1], 'y': [0, 1]},
            value=exklp,
            mode="gauge+number",
            title={'text': "Exklusivanteil", 'font': {
                'size': 40, 'family': 'voestalpine'}},
            #delta = {'reference': 50},
            number={'suffix': "%"},
            gauge={
                "bar": {"color": "red"},
                'axis': {'range': [None, 100]}
            }))

        # Zeigerdiagramm (Universalanteil) anzeigen
        with col111:
            st.plotly_chart(fig_uni, use_container_width=True, height=500)

        # Zeigerdiagramm (Exklusivanteil) anzeigen
        with col113:
            st.plotly_chart(fig_exkl, use_container_width=True)

        # Häufigsten Bauteile ermitteln
        n = 5
        topl = df['Objektsparte'].value_counts().index.tolist()[:n]
        topl.insert(0, "test")
        data = {"Komponenten": topl}
        dftop = pd.DataFrame(data)
        dftop = dftop.iloc[1:, :]

        # Häufigesten Bauteile mittels Dataframe visualisieren
        with col112:
            st.markdown(
                '<p style="text-align: center;color: #0082B4;font-size:30px">Häufigsten<br>Komponenten</p>', unsafe_allow_html=True)
            st.write(dftop.astype("object"), width=120)

# Option Portfolio
if selected_hor == "Portfolio":

    # Überschrift einfügen
    with col1:
        st.markdown("""
            <style>
            .big-font {
            font-size:50px !important;
            color: #0082B4
            }
            </style>
            """, unsafe_allow_html=True)
        st.markdown('<p class="big-font">Portfolio</p>',
                    unsafe_allow_html=True)

    # Railway Systems Logo einfügen
    with col2:
        st.image('https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

    # Datei-Uploader einfügen
    with st.expander("Upload"):
        iFiles = st.file_uploader(
            "", accept_multiple_files=True, type=["xlsm"])

    # Dateien nach Benennung sortieren
    iFiles.sort(key=lambda x: x.name.split("_")[2])

    # Produktnamen filtern
    iNames = []
    for i in iFiles:
        iNames.append(i.name.split("_")[0]+i.name.split("_")[2])

    # Produktnamen in Liste umwandeln
    iNames = list(dict.fromkeys(iNames))

    # Produktauswahl
    iFile = st.selectbox("Analyse-File Auswahl", iNames, key="PORT-V")

    # Dateien zuweisen
    files = []
    for i in iFiles:
        if i.name.split("_")[0]+i.name.split("_")[2] == iFile:
            files.append(i)

    # Sidebar erstellen
    with st.sidebar:
        # VoestAlpine Logo einfügen
        st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')
        # Sidebar Optionsmenü erstellen
        selected_side = option_menu(
            menu_title="Materialgruppe",
            options=("Allgemein", "Fertigungsteile",
                     "Zukaufteile", "Normteile"),
            icons=("caret-right-fill", "caret-right-fill",
                   "caret-right-fill", "caret-right-fill"),
            default_index=0,
            menu_icon="cast",
            styles={
                "container": {"padding": "0!important", "background-color": "#262730"},
                "icon": {"color": "#FAFAFA"},
                "nav-link": {"font_family": "voestalpine", "text-align": "left", "--hover-color": "#A5A5A5"},
                "nav-link-selected": {"background-color": "#0082B4"},
            }
        )

    # Dateien sortieren
    files.sort(key=lambda x: x.name.split("_")[1])

    # Bei Produktauswahl
    if files is not None:

        try:
            # Selectbox einfügen
            file = None

            # jahrauswahl
            option1 = {
                "dataset": {},
                "tooltip": {
                    "show": False,
                },
                "series": [
                    {
                        "type": 'liquidFill',
                        "name": "2018",
                        "center": ["20%", "50%"],
                        "data":[1],
                        "itemStyle": {
                            "color": "white",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "formatter": "{a}",
                            "insideColor": "#0082B4",
                            "fontSize": "20"
                        },
                        "radius": "90%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "grey"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": '#0082B4',
                                "shadowBlur": 0,
                                "shadowColor": 'green'
                            }}},
                    {
                        "type": 'liquidFill',
                        "name": "2019",
                        "center": ["40%", "50%"],
                        "data":[1],
                        "itemStyle": {
                            "color": "white",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "formatter": "{a}",
                            "insideColor": "#0082B4",
                            "fontSize": "20"
                        },
                        "radius": "90%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "grey"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": '#0082B4',
                                "shadowBlur": 0,
                                "shadowColor": 'green'
                            }}},
                    {
                        "type": 'liquidFill',
                        "name": "2020",
                        "center": ["60%", "50%"],
                        "data":[1],
                        "itemStyle": {
                            "color": "white",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "formatter": "{a}",
                            "insideColor": "#0082B4",
                            "fontSize": "20"
                        },
                        "radius": "90%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "grey"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": '#0082B4',
                                "shadowBlur": 0,
                                "shadowColor": 'green'
                            }}},
                    {
                        "type": 'liquidFill',
                        "name": "2021",
                        "center": ["80%", "50%"],
                        "data":[1],
                        "itemStyle": {
                            "color": "white",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "formatter": "{a}",
                            "insideColor": "#0082B4",
                            "fontSize": "20"
                        },
                        "radius": "90%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "grey"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": '#0082B4',
                                "shadowBlur": 0,
                                "shadowColor": 'green'
                            }}},
                ]
            }

            # Click-Events
            events1 = {
                "click": "function(params) { return params.seriesName }"}

            # Click-Output
            output1 = st_echarts(option1, height=100, events=events1, key="1")

            # Jahrauswahl zuweisen
            for i in files:
                if i.name.split("_")[1] == output1:
                    file = i

            # Jahrauswahl formatieren
            Jahr1 = f"""<style>p.a {{font:50px voestalpine;color: #0082B4;text-align: center}}</style><p class="a">von {output1}</p>"""

            # Jahrauswahl darstellen
            if output1 is not None:
                st.markdown(Jahr1, unsafe_allow_html=True)

            # Portfolio-Auswertung einlesen
            df = pd.read_excel(file,
                               engine="openpyxl",
                               sheet_name="Portfolio_Auswertung",
                               na_filter=True)

            # Portfolio-Daten einlesen
            dfp = pd.read_excel(file,
                                engine="openpyxl",
                                sheet_name="Portfolio",
                                na_filter=True)

            # NA-Zeilen entfernen
            dfp.dropna(inplace=True)

            # Spalte in Typ String umwandeln
            dfp = dfp.astype({'MatNr.': 'string'})

            # Nach Materialart filtern
            if selected_side == "Fertigungsteile":
                dfp = dfp.loc[(dfp['Materialart'] == "FERT") |
                              (dfp['Materialart'] == "60FE")]

            if selected_side == "Zukaufteile":
                dfp = dfp.loc[(dfp['Materialart'] == "HIBE") |
                              (dfp['Materialart'] == "60HI")]

            if selected_side == "Normteile":
                dfp = dfp.loc[dfp['Materialart'] == "NORM"]

            # Spalten definieren
            col111, col112, col113, col114 = st.columns((7, 6, 6, 6))
            col11, col12 = st.columns((2, 5))

            # Spaltenbeschriftung
            with col112:
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:35px">Klasse A</p>', unsafe_allow_html=True)

            with col113:
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:35px">Klasse B</p>', unsafe_allow_html=True)

            with col114:
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:35px">Klasse C</p>', unsafe_allow_html=True)

            # Zeilenbeschriftung
            with col11:
                st. write("")
                st. write("")
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">Top Effizienz</p>', unsafe_allow_html=True)
                st.markdown(
                    '<p style="text-align: center;color: #FFF;font-size:26px">66 - 100%</p>', unsafe_allow_html=True)
                st. write("")
                st. write("")
                st. write("")
                st. write("")
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">Medium Effizienz</p>', unsafe_allow_html=True)
                st.markdown(
                    '<p style="text-align: center;color: #FFF;font-size:26px">33 - 66%</p>', unsafe_allow_html=True)
                st. write("")
                st. write("")
                st. write("")
                st. write("")
                st.markdown(
                    '<p style="text-align: center;color: #0082B4;font-size:30px">Low Effizienz</p>', unsafe_allow_html=True)
                st.markdown(
                    '<p style="text-align: center;color: #FFF;font-size:26px">0 - 33%</p>', unsafe_allow_html=True)
                st. write("")
                st. write("")
                st. write("")

            # Portfolio
            with col12:
                option = {
                    "dataset": {},
                    "tooltip": {
                        "show": True,
                    },
                    "series": [
                        {
                            "type": 'liquidFill',
                            "name": "Top-C",
                            "center": ["85%", "13%"],
                            "data":[df.at[8, "Anteil"]],
                            "itemStyle": {
                                "color": "#0082B4",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "color": "#FFF",
                            },
                            "radius": "25%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "#0E1117"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": 'orange',
                                    "shadowBlur": 0,
                                    "shadowColor": '#0082B4'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "Medium-C",
                            "center": ["85%", "42%"],
                            "data":[df.at[7, "Anteil"]],
                            "itemStyle": {
                                "color": "#0082B4",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "color": "#FFF",
                            },
                            "radius": "25%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "#0E1117"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": 'red',
                                    "shadowBlur": 0,
                                    "shadowColor": '#0082B4'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "Low-C",
                            "center": ["85%", "71%"],
                            "data":[df.at[6, "Anteil"]],
                            "itemStyle": {
                                "color": "#0082B4",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "color": "#FFF",
                            },
                            "radius": "25%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "#0E1117"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": 'red',
                                    "shadowBlur": 0,
                                    "shadowColor": '#0082B4'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "Top-B",
                            "center": ["50%", "13%"],
                            "data":[df.at[5, "Anteil"]],
                            "itemStyle": {
                                "color": "#0082B4",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "color": "#FFF",
                            },
                            "radius": "25%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "#0E1117"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": 'yellow',
                                    "shadowBlur": 0,
                                    "shadowColor": '#0082B4'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "Medium-B",
                            "center": ["50%", "42%"],
                            "data":[df.at[4, "Anteil"]],
                            "itemStyle": {
                                "color": "#0082B4",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "color": "#FFF",
                            },
                            "radius": "25%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "#0E1117"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": 'orange',
                                    "shadowBlur": 0,
                                    "shadowColor": '#0082B4'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "Low-B",
                            "center": ["50%", "71%"],
                            "data":[df.at[3, "Anteil"]],
                            "itemStyle": {
                                "color": "#0082B4",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "color": "#FFF",
                            },
                            "radius": "25%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "#0E1117"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": 'red',
                                    "shadowBlur": 0,
                                    "shadowColor": '#0082B4'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "Top-A",
                            "center": ["15%", "13%"],
                            "data":[df.at[2, "Anteil"]],
                            "itemStyle": {
                                "color": "#0082B4",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "color": "#FFF",
                            },
                            "radius": "25%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "#0E1117"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": '#0eff00',
                                    "shadowBlur": 0,
                                    "shadowColor": '#0082B4'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "Medium-A",
                            "center": ["15%", "42%"],
                            "data":[df.at[1, "Anteil"]],
                            "itemStyle": {
                                "color": "#0082B4",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "color": "#FFF",
                            },
                            "radius": "25%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "#0E1117"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": 'yellow',
                                    "shadowBlur": 0,
                                    "shadowColor": '#0082B4'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "Low-A",
                            "center": ["15%", "71%"],
                            "data":[df.at[0, "Anteil"]],
                            "itemStyle": {
                                "color": "#0082B4",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "color": "#FFF",
                            },
                            "radius": "25%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "#0E1117"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": 'orange',
                                    "shadowBlur": 0,
                                    "shadowColor": '#0082B4'
                                }}},
                        {
                            "type": 'liquidFill',
                            "name": "Reset",
                            "center": ["50%", "93%"],
                            "data":[1],
                            "itemStyle": {
                                "color": "white",
                                "shadowBlur": 0
                            },
                            "amplitude": 0,
                            "label": {
                                "formatter": "{a}",
                                "insideColor": "#0082B4",
                                "fontSize": "20"
                            },
                            "radius": "13%",
                            "waveAnimation": 0,
                            "shape": 'roundRect',
                            "backgroundStyle": {
                                "color": "grey"
                            },
                            "outline": {
                                "borderDistance": 5,
                                "itemStyle": {
                                    "borderWidth": 5,
                                    "borderColor": '#0082B4',
                                    "shadowBlur": 0,
                                    "shadowColor": 'green'
                                }}}
                    ]
                }

                events = {
                    "click": "function(params) { return params.seriesName }"}

                output = st_echarts(option, height=730, events=events)

            # Bei Reset
            if output == "Reset":
                output = None

            # Datenmaske
            mask = None

            # Klasse filtern
            frame = dfp.loc[dfp['Klasse'] == output]

            # NA-Zeilen entfernen
            frame.dropna(subset=["Materialart"], inplace=True)

            # Spalten filtern
            f = frame.astype(
                str).loc[:, ["MatNr.", "Objektkurztext", "Materialart", "Warengruppe"]]

            gb = GridOptionsBuilder.from_dataframe(f)

            # Portfolio
            grid_options = {
                "defaultColDef": {
                    "minWidth": 5,
                    "editable": False,
                    "filter": True,
                    "resizable": True,
                    "sortable": True
                },
                "suppressFieldDotNotation": True,
                "columnDefs": [
                    {
                        "headerName": "MatNr.",
                        "field": "MatNr.",
                        "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "20px"},
                        "width": 270,
                        "type": []
                    },
                    {
                        "headerName": "Objektkurztext",
                        "field": "Objektkurztext",
                        "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "15px"},
                        "width": 600,
                        "type": []
                    },
                    {
                        "headerName": "Materialart",
                        "field": "Materialart",
                        "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "20px"},
                        "type": []
                    },
                    {
                        "headerName": "Warengruppe",
                        "field": "Warengruppe",
                        "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "20px"},
                        "type": []
                    },
                ]
            }

            # Info-Box
            if output != None:

                if output == "Top-A":
                    st.info(
                        "Komponenten dieser Kategorie werden in vielen Varianten eingesetzt, die einen hohen Umsatz generieren.")

                if output == "Top-B":
                    st.info(
                        "Komponenten dieser Kategorie werden in vielen Varianten eingesetzt, sind jedoch nicht umsatzentscheidend.")

                if output == "Top-C":
                    st.info(
                        "Komponenten dieser Kategorie werden in vielen Varianten eingesetzt, bringen jedoch keinen Umsatz.")

                if output == "Medium-A":
                    st.info(
                        "Komponenten dieser Kategorie weisen eine mittlere Effizienz auf, generieren jedoch einen hohen Umsatz.")

                if output == "Medium-B":
                    st.info(
                        "Komponenten dieser Kategorie weisen eine mittlere Effizienz auf, sind jedoch nicht umsatzentscheidend.")

                if output == "Medium-C":
                    st.info(
                        "Komponenten dieser Kategorie weisen eine mittlere Effizienz auf, bringen jedoch keinen Umsatz.")

                if output == "Low-A":
                    st.info(
                        "Komponenten dieser Kategorie werden nur in wenigen Varianten verbaut, welche jedoch einen hohen Anteil am Umsatz haben.")

                if output == "Low-B":
                    st.info(
                        "Komponenten dieser Kategorie werden nur in wenigen Varianten verbaut und generieren keinen wesentlichen Umsatz.")

                if output == "Low-C":
                    st.info(
                        "Komponenten dieser Kategorie werden nur in wenigen Varianten verbaut und generieren keinen Umsatz.")

            # Spalten definieren
            colp1, colp2 = st.columns((1, 1))

            # Suchfunktionen
            with colp1:
                sucheM = st.text_input(
                    "Materialnummer suchen...", placeholder="z.B. 73100000001A")
            with colp2:
                sucheO = st.text_input(
                    "Objektkurztext suchen...", placeholder="z.B. Angriffslappen")

            # Sucheinstellungen
            if output is not None:
                if sucheM and sucheO != "":
                    mask1 = f["MatNr."].str.contains(
                        sucheM, case=False, na=False)
                    mask2 = f["Objektkurztext"].str.contains(
                        sucheO, case=False, na=False)
                    mask = mask1 & mask2

                else:
                    if sucheM != "":
                        mask = f["MatNr."].str.contains(
                            sucheM, case=False, na=False)
                    if sucheO != "":
                        mask = f["Objektkurztext"].str.contains(
                            sucheO, case=False, na=False)

                if mask is not None:
                    AgGrid(f[mask], gridOptions=grid_options, theme="streamlit",
                           height=600, fit_columns_on_grid_load=True, editable=False)

                    buffer = io.BytesIO()

                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        f[mask].to_excel(writer)

                    col1db, col2db, col3db = st.columns((1, 1, 1))
                    with col2db:
                        st.download_button(
                            label="Download Auswertung",
                            data=buffer,
                            file_name=file.name.split(
                                "_")[0]+file.name.split("_")[1]+"_"+output+".xlsx",
                            mime="application/vnd.ms-excel"
                        )
                else:
                    AgGrid(f, gridOptions=grid_options, theme="streamlit",
                           height=600, fit_columns_on_grid_load=True, editable=False)

                    buffer = io.BytesIO()

                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        f.to_excel(writer)

                    col1db, col2db, col3db = st.columns((1, 1, 1))
                    with col2db:
                        st.download_button(
                            label="Download Auswertung",
                            data=buffer,
                            file_name=file.name.split(
                                "_")[0]+file.name.split("_")[1]+"_"+output+".xlsx",
                            mime="application/vnd.ms-excel"
                        )
            else:
                if sucheM and sucheO != "":
                    mask1 = dfp["MatNr."].str.contains(
                        sucheM, case=False, na=False)
                    mask2 = dfp["Objektkurztext"].str.contains(
                        sucheO, case=False, na=False)
                    mask = mask1 & mask2

                else:
                    if sucheM != "":
                        mask = dfp["MatNr."].str.contains(
                            sucheM, case=False, na=False)
                    if sucheO != "":
                        mask = dfp["Objektkurztext"].str.contains(
                            sucheO, case=False, na=False)

                if mask is not None:
                    AgGrid(dfp[mask], gridOptions=grid_options, theme="streamlit",
                           height=600, fit_columns_on_grid_load=True, editable=False)

                    buffer = io.BytesIO()

                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        dfp[mask].to_excel(writer)

                    col1db, col2db, col3db = st.columns((1, 1, 1))
                    with col2db:
                        st.download_button(
                            label="Download Auswertung",
                            data=buffer,
                            file_name=file.name.split(
                                "_")[0]+file.name.split("_")[1]+".xlsx",
                            mime="application/vnd.ms-excel"
                        )

                else:
                    AgGrid(dfp, gridOptions=grid_options, theme="streamlit",
                           height=600, fit_columns_on_grid_load=True, editable=False)

                    buffer = io.BytesIO()

                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        dfp.to_excel(writer)

                    col1db, col2db, col3db = st.columns((1, 1, 1))
                    with col2db:
                        st.download_button(
                            label="Download Auswertung",
                            data=buffer,
                            file_name=file.name.split(
                                "_")[0]+file.name.split("_")[1]+".xlsx",
                            mime="application/vnd.ms-excel"
                        )

        # Bei fehlenden Input:
        except ValueError:
            st.markdown(
                '<p style="text-align: center;color: #FFF;font-size:20px">Bitte uploaden Sie die Analyse-Dateien!</p>', unsafe_allow_html=True)

# Option Stückliste
if selected_hor == "Stückliste":

    # Überschrift einfügen
    with col1:
        st.markdown("""
            <style>
            .big-font {
            font-size:50px !important;
            color: #0082B4
            }
            </style>
            """, unsafe_allow_html=True)
        st.markdown('<p class="big-font">Stückliste</p>',
                    unsafe_allow_html=True)

    # Railway Systems Logo einfügen
    with col2:
        st.image('https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

    # Datei-Uploader einfügen
    with st.expander("Upload"):
        iFiles = st.file_uploader(
            "", accept_multiple_files=True, type=["xlsm"])

    # Dateien nach Benennung sortieren
    iFiles.sort(key=lambda x: x.name.split("_")[2])

    # Produktnamen filtern
    iNames = []
    for i in iFiles:
        iNames.append(i.name.split("_")[0]+i.name.split("_")[2])

    # Produktnamen in Liste umwandeln
    iNames = list(dict.fromkeys(iNames))

    # Produktauswahl
    iFile = st.selectbox("Analyse-File Auswahl", iNames)

    # Produktauswahl zuweisen
    files = []
    for i in iFiles:
        if i.name.split("_")[0]+i.name.split("_")[2] == iFile:
            files.append(i)

    # Dateien sortieren
    files.sort(key=lambda x: x.name.split("_")[1])

    # Bei Dateiauswahl
    if files is not None:

        try:
            # Selectbox einfügen
            file = None

            # Jahrauswahl
            option1 = {
                "dataset": {},
                "tooltip": {
                    "show": False,
                },
                "series": [
                    {
                        "type": 'liquidFill',
                        "name": "2018",
                        "center": ["20%", "50%"],
                        "data":[1],
                        "itemStyle": {
                            "color": "white",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "formatter": "{a}",
                            "insideColor": "#0082B4",
                            "fontSize": "20"
                        },
                        "radius": "90%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "grey"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": '#0082B4',
                                "shadowBlur": 0,
                                "shadowColor": 'green'
                            }}},
                    {
                        "type": 'liquidFill',
                        "name": "2019",
                        "center": ["40%", "50%"],
                        "data":[1],
                        "itemStyle": {
                            "color": "white",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "formatter": "{a}",
                            "insideColor": "#0082B4",
                            "fontSize": "20"
                        },
                        "radius": "90%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "grey"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": '#0082B4',
                                "shadowBlur": 0,
                                "shadowColor": 'green'
                            }}},
                    {
                        "type": 'liquidFill',
                        "name": "2020",
                        "center": ["60%", "50%"],
                        "data":[1],
                        "itemStyle": {
                            "color": "white",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "formatter": "{a}",
                            "insideColor": "#0082B4",
                            "fontSize": "20"
                        },
                        "radius": "90%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "grey"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": '#0082B4',
                                "shadowBlur": 0,
                                "shadowColor": 'green'
                            }}},
                    {
                        "type": 'liquidFill',
                        "name": "2021",
                        "center": ["80%", "50%"],
                        "data":[1],
                        "itemStyle": {
                            "color": "white",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "formatter": "{a}",
                            "insideColor": "#0082B4",
                            "fontSize": "20"
                        },
                        "radius": "90%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "grey"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": '#0082B4',
                                "shadowBlur": 0,
                                "shadowColor": 'green'
                            }}},
                ]
            }

            # Click-Events
            events1 = {
                "click": "function(params) { return params.seriesName }"}

            # Click-Output
            output1 = st_echarts(option1, height=100, events=events1, key="1")

            # Click-Output zuweisen
            for i in files:
                if i.name.split("_")[1] == output1:
                    file = i

            # Jahrauswahl formatieren
            Jahr1 = f"""<style>p.a {{font:50px voestalpine;color: #0082B4;text-align: center}}</style><p class="a">von {output1}</p>"""

            # Jahrauswahl darstellen
            if output1 is not None:
                st.markdown(Jahr1, unsafe_allow_html=True)

            # Eingabe einlesen
            df = pd.read_excel(file,
                               engine="openpyxl",
                               sheet_name="Eingabe",
                               na_filter=True,
                               usecols="A")

            # Varianten zuweisen
            var = []
            var = df["Materialnummer"].values.tolist()
            var.sort()

            # Variantenauswahl
            varA = st.selectbox("Varianten Auswahl", var)

            # Variantenauswahl einlesen
            dfA = pd.read_excel(file,
                                engine="openpyxl",
                                sheet_name=varA + ".xlsx",
                                na_filter=True)

            # Komponenten einlesen
            dfK = pd.read_excel(file,
                                engine="openpyxl",
                                sheet_name="Komponenten",
                                na_filter=True)

            # Komponentennummern in Liste umwandeln
            K = []
            K = dfA["Komponentennummer"].values.tolist()

            # Initialisierung
            eff = []
            effT = []
            c = 0

            # Komponenten suchen
            for j in K:
                c = c+1
                for i in range(len(dfK.Komponentennummer)):
                    if dfK.Komponentennummer[i] == j:
                        eff.append(float(dfK.Effizienz[i]*100))
                if len(eff) < c:
                    eff.append(0.0)

            # Effizienzspalte zuweisen
            dfE = pd.DataFrame(eff, columns=["Effizienz"])

            # Spalte an Dataframe anhängen
            dfG = pd.concat([dfA, dfE], axis=1)

            # Dataframe filtern
            dfG = dfG.query("Warengruppe !='P24401'")

            # Dataframe sorieren
            dfG = dfG.sort_values('Effizienz', ascending=True)

            # NA-Zeilen entfernen
            dfG.dropna(subset=["Objektkurztext"], inplace=True)

            # In Prozentwerte umwandeln
            dfG['Effizienz'] = dfG['Effizienz'].map('{:,.1f}%'.format)

            gb = GridOptionsBuilder.from_dataframe(dfG)

            # Dataframe-Einstellungen
            grid_options = {
                "defaultColDef": {
                    "minWidth": 5,
                    "editable": False,
                    "filter": True,
                    "resizable": True,
                    "sortable": True
                },
                "columnDefs": [
                    {
                        "headerName": "Komponentennummer",
                        "field": "Komponentennummer",
                        "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "18px"},
                        "width": 100,
                        "type": []
                    },
                    {
                        "headerName": "Objektkurztext",
                        "field": "Objektkurztext",
                        "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "18px"},
                        "width": 250,
                        "type": []
                    },
                    {
                        "headerName": "Materialart",
                        "field": "Materialart",
                        "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "18px"},
                        "width": 40,
                        "type": []
                    },
                    {
                        "headerName": "Warengruppe",
                        "field": "Warengruppe",
                        "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "18px"},
                        "width": 60,
                        "type": []
                    },
                    {
                        "headerName": "Effizienz",
                        "field": "Effizienz",
                        "cellStyle": {'background-color': "#0E1117", "font-family": "voestalpine", "font-size": "18px"},
                        "width": 60,
                        "type": []
                    },
                ]
            }

            # Dataframe anzeigen
            AgGrid(dfG, gridOptions=grid_options, theme="streamlit",
                   height=400, fit_columns_on_grid_load=True, editable=False)

            # Geladene Daten einlesen
            buffer = io.BytesIO()

            # Dataframe in Excel umwandeln
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                dfG.to_excel(writer)

            # Spalten definieren
            col1db, col2db, col3db = st.columns((1, 1, 1))
            with col2db:
                # Download-Button einfügen
                st.download_button(
                    label="Download Auswertung",
                    data=buffer,
                    file_name=varA+"_Analyse.xlsx",
                    mime="application/vnd.ms-excel"
                )

        # Bei fehlenden Input:
        except ValueError:
            st.markdown(
                '<p style="text-align: center;color: #FFF;font-size:20px">Bitte uploaden Sie die Analyse-Dateien!</p>', unsafe_allow_html=True)
