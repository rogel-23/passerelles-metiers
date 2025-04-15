import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
from datetime import datetime

# ------------------------------
# 🔐 Sécurité : accès par mot de passe
# ------------------------------
CORRECT_PASSWORD = "Passerelle2025"

password = st.text_input("🔒 Veuillez entrer le mot de passe pour accéder à l'outil :", type="password")

if password != CORRECT_PASSWORD:
    st.warning("Mot de passe incorrect ou manquant. Veuillez entrer le bon mot de passe.")
    st.stop()

# ------------------------------
# TITRE & UPLOAD
# ------------------------------
st.title("🧭 Outil de passerelles métiers")

st.markdown("""
---
👋 Bienvenue dans l'outil de passerelles métiers !

Cet outil vous permet d’identifier les **passerelles métiers possibles** entre les métiers d’un client (ex : secteur pharmaceutique) et ceux du référentiel ROME, sur la base des **macro-compétences partagées**.

---

### 🧩 Étapes à suivre :
1. **Charger les deux fichiers Excel**.
2. **Choisir le type de passerelle** (entrante ou sortante).
3. **Sélectionner les dimensions de compétences** à prendre en compte ainsi que leur **pondération**.
4. **Filtrer par secteur** si besoin, puis **choisir un métier de départ**.
5. 📊 Obtenez les passerelles les plus proches et **téléchargez les résultats**.
6. (Facultatif) 📦 Générez **l’intégralité des passerelles** sans aucun filtre.

---
""")

st.markdown("###\n**📚 1. Charger le fichier des compétences ROME (MACRO-COMPETENCES ROME.xlsx)**")
fichier_competences = st.file_uploader("", type="xlsx", key="competences")
st.markdown("###\n**🏢 2. Charger le fichier des métiers client**")
st.markdown("""
<small>ℹ️ Le fichier métiers client doit contenir **une colonne intitulée `Code ROME`**, avec un code ROME par ligne (ex : M1805).<br>
Autres colonnes (intitulé, descriptions…) facultatives.</small>
""", unsafe_allow_html=True)
fichier_client = st.file_uploader("", type="xlsx", key="client")

if fichier_competences and fichier_client:
    # ------------------------------
    # CHOIX DU MODE ET DES OPTIONS
    # ------------------------------
    st.markdown("###\n**🔍 3. Type de passerelle**")
    mode = st.radio("", ["Passerelle entrante", "Passerelle sortante"])

    st.markdown("###\n**🎯 4. Catégories de compétences**")
    col1, col2, col3 = st.columns(3)
    with col1:
        avec_savoir_faire = st.checkbox("Savoir-faire", value=True)
    with col2:
        avec_savoir_etre = st.checkbox("Savoir-être professionnels", value=True)
    with col3:
        avec_savoirs = st.checkbox("Savoirs", value=True)

    st.markdown("###\n**⚖️ 5. Pondération des catégories de compétences (total = 100%)**")
    col_w1, col_w2, col_w3 = st.columns(3)

    with col_w1:
        poids_sf = st.number_input("🛠️ Savoir-faire (%)", min_value=0, max_value=100, value=20, step=5, disabled=not avec_savoir_faire)

    with col_w2:
        poids_se = st.number_input("🤝 Savoir-être (%)", min_value=0, max_value=100, value=20, step=5, disabled=not avec_savoir_etre)

    with col_w3:
        poids_savoirs = st.number_input("📚 Savoirs (%)", min_value=0, max_value=100, value=60, step=5, disabled=not avec_savoirs)

    # Calcul dynamique selon cases cochées
    total_pondere = 0
    if avec_savoir_faire:
        total_pondere += poids_sf
    if avec_savoir_etre:
        total_pondere += poids_se
    if avec_savoirs:
        total_pondere += poids_savoirs

    if total_pondere != 100:
        st.error("❌ La somme des pondérations doit être égale à 100% pour les catégories sélectionnées.")
        st.stop()

    # Liste des catégories sélectionnées
    categories_selectionnees = []
    if avec_savoir_faire:
        categories_selectionnees.append("Savoir-faire")
    if avec_savoir_etre:
        categories_selectionnees.append("Savoir-être professionnels")
    if avec_savoirs:
        categories_selectionnees.append("Savoirs")
    
    if not categories_selectionnees:
        st.warning("⚠️ Veuillez sélectionner au moins une catégorie de compétence (savoir-faire, savoir-être professionnels ou savoirs).")
        st.stop() 

    # Chargement unique depuis l'onglet centralisé
    df_comp_brut = pd.read_excel(fichier_competences, sheet_name="Macro-Compétences")
    df_comp = df_comp_brut[df_comp_brut["Catégorie"].isin(categories_selectionnees)].copy()
    df_comp = df_comp.dropna(subset=["Code Métier", "Intitulé", "Macro Compétence"])  # Nettoyage

    # Chargement des métiers client
    df_client = pd.read_excel(fichier_client)
    codes_client = df_client["Code ROME"].dropna().unique()

    # Définir métiers de départ et d'arrivée selon le mode
    if mode == "Passerelle entrante":
        df_depart = df_comp.copy()  # Tous les métiers (ROME + client)
        df_arrivee = df_comp[df_comp["Code Métier"].isin(codes_client)]
    else:
        df_depart = df_comp[df_comp["Code Métier"].isin(codes_client)]
        df_arrivee = df_comp[~df_comp["Code Métier"].isin(codes_client)]


    # Liste des métiers de départ disponibles
    metiers_depart = df_depart[["Code Métier", "Intitulé"]].drop_duplicates().sort_values("Intitulé")

    # Dictionnaire de correspondance lettre → secteur
    secteurs = {
        "A": "Agriculture et Pêche, Espaces naturels et Espaces verts, Soins aux animaux",
        "B": "Arts et Façonnage d'ouvrages d'art",
        "C": "Banque, Assurance, Immobilier",
        "D": "Commerce, Vente et Grande distribution",
        "E": "Communication, Média et Multimédia",
        "F": "Construction, Bâtiment et Travaux publics",
        "G": "Hôtellerie-Restauration, Tourisme, Loisirs et Animation",
        "H": "Industrie",
        "I": "Installation et Maintenance",
        "J": "Santé",
        "K": "Services à la personne et à la collectivité",
        "L": "Spectacle",
        "M": "Support à l'entreprise",
        "N": "Transport et Logistique"
    }

    # Lettres présentes dans les métiers de départ
    lettres_disponibles = df_depart["Code Métier"].str[0].unique()
    secteurs_disponibles = {lettre: secteurs[lettre] for lettre in lettres_disponibles if lettre in secteurs}

    # Construction des options de filtre secteur
    options_secteurs = ["Tous les secteurs"] + [f"{lettre} - {secteurs[lettre]}" for lettre in sorted(secteurs_disponibles)]

    # Initialisation du filtre secteur en session_state
    if "secteur_selectionne" not in st.session_state:
        st.session_state["secteur_selectionne"] = "Tous les secteurs"

    # Menu déroulant secteur
    st.markdown("###\n**🗂️ 6. Secteur d'activité**")
    secteur_selectionne = st.selectbox(
        "",
        options=options_secteurs,
        index=options_secteurs.index(st.session_state["secteur_selectionne"]),
        key="secteur_selectionne"
    )

    # Filtrage des métiers de départ si un secteur est sélectionné
    if secteur_selectionne == "Tous les secteurs":
        metiers_filtrés = metiers_depart.copy()
    else:
        lettre_selectionnee = secteur_selectionne.split(" - ")[0]
        metiers_filtrés = metiers_depart[metiers_depart["Code Métier"].str.startswith(lettre_selectionnee)]

    # Liste des métiers de départ disponibles
    st.markdown("###\n**👤 7. Métier de départ**")
    metiers_filtrés["Affichage"] = metiers_filtrés["Code Métier"] + " - " + metiers_filtrés["Intitulé"]
    choix_affichage = st.selectbox("", options=metiers_filtrés["Affichage"].tolist())
    code_selectionne = metiers_filtrés[metiers_filtrés["Affichage"] == choix_affichage]["Code Métier"].values[0]
    choix_metier = metiers_filtrés[metiers_filtrés["Code Métier"] == code_selectionne]["Intitulé"].values[0]


    # Code métier sélectionné
    code_selectionne = metiers_depart[metiers_depart["Intitulé"] == choix_metier]["Code Métier"].values[0]

    # Macro-compétences du métier sélectionné
    competences_selection = set(df_depart[df_depart["Code Métier"] == code_selectionne]["Macro Compétence"].dropna())

    # Calcul des similarités avec les métiers d'arrivée
    lignes_resultats = []
    lettre_depart = code_selectionne[0]

    for code_metier, groupe in df_arrivee.groupby("Code Métier"):
        if code_metier == code_selectionne:
            continue  # On saute si c'est le même métier
        intitule = groupe["Intitulé"].iloc[0]
        competences_metier = set(groupe["Macro Compétence"].dropna())
        intersection = competences_selection & competences_metier
        score = len(intersection)
        if score > 0:
            for comp in intersection:
                # On récupère la catégorie depuis le df_arrivee
                categorie = groupe[groupe["Macro Compétence"] == comp]["Catégorie"].iloc[0]
                    # Appliquer le poids en fonction de la catégorie
                if categorie == "Savoir-faire":
                    poids = poids_sf
                elif categorie == "Savoir-être professionnels":
                    poids = poids_se
                elif categorie == "Savoirs":
                    poids = poids_savoirs
                else:
                    poids = 0  # sécurité

                bonus = 1.25 if code_metier.startswith(lettre_depart) else 1
                score_pondere_final = (score * poids / 100) * bonus
                
                lignes_resultats.append({
                    "Code Métier": code_metier,
                    "Intitulé": intitule,
                    "Nb de passerelles communes": score,
                    "Score pondéré": score * poids / 100,
                    "Catégorie": categorie,
                    "Compétence commune": comp
                })

    if lignes_resultats:
        df_resultats_complets = pd.DataFrame(lignes_resultats)

        # Affichage top 20 pour l'écran
        top_metiers = (
            df_resultats_complets.groupby(["Code Métier", "Intitulé"])
            .agg({
                "Score pondéré": "sum",
                "Compétence commune": "count"
            })
            .reset_index()
            .rename(columns={
                "Score pondéré": "Score pondéré total",
                "Compétence commune": "Nombre de compétences partagées"
            })
            .sort_values("Score pondéré total", ascending=False)
            .head(20)
        )
        st.markdown("###\n### 🌟 Top 20 des passerelles proposées :")
        st.dataframe(top_metiers)

        # 🔍 Graphique des scores (bar chart)
        # 🔁 Créer un tableau croisé avec Score pondéré par catégorie
        df_pivot = df_resultats_complets.pivot_table(
            index=["Code Métier", "Intitulé"],
            columns="Catégorie",
            values="Score pondéré",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        # 🔁 Recalcul du score total pour tri
        df_pivot["Score total"] = df_pivot[categories_selectionnees].sum(axis=1)
        df_pivot = df_pivot.sort_values("Score total", ascending=False).head(20)

        # 🔍 Graphique empilé
        st.markdown("### 📊 Répartition des scores pondérés par type de compétence")
        fig, ax = plt.subplots(figsize=(8, 6))

        bottom = None
        labels = df_pivot["Intitulé"]

        # Colorer chaque barre selon la catégorie
        for cat in categories_selectionnees:
            ax.barh(labels, df_pivot[cat], left=bottom, label=cat)
            if bottom is None:
                bottom = df_pivot[cat].copy()
            else:
                bottom += df_pivot[cat]

        ax.invert_yaxis()
        ax.set_xlabel("Score pondéré")
        ax.set_title("Top 20 métiers – scores par type de compétence")
        ax.legend(title="Catégorie")

        st.pyplot(fig)

        # Filtres et formats Excel
        df_filtré = df_resultats_complets.copy()
        categories_str = ", ".join(categories_selectionnees)

        # Ajouter le score pondéré total dans df_filtré
        df_scores = (
            df_filtré.groupby(["Code Métier", "Intitulé"])["Score pondéré"]
            .sum()
            .reset_index()
            .rename(columns={"Score pondéré": "Score pondéré total"})
        )

        # Merge avec le fichier filtré ligne par ligne
        df_filtré = df_filtré.merge(df_scores, on=["Code Métier", "Intitulé"], how="left")

        # Réorganiser les colonnes et supprimer "Score pondéré"
        colonnes_ordre = [
            "Code Métier",
            "Intitulé",
            "Score pondéré total",
            "Nb de passerelles communes",
            "Catégorie",
            "Compétence commune"
        ]
        df_filtré = df_filtré[colonnes_ordre]

        # Trier par Score pondéré total décroissant
        df_filtré = df_filtré.sort_values(
            by=["Score pondéré total", "Code Métier", "Intitulé"],
            ascending=[False, True, True]
        )

        def exporter_excel(df, titre_feuille, nom_fichier):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet(titre_feuille)
                writer.sheets[titre_feuille] = worksheet

                format_titre = workbook.add_format({'bold': True, 'font_size': 14})
                format_soustitre = workbook.add_format({'italic': True})

                # Écriture du titre et des dimensions choisies
                worksheet.write("A1", f"Métier de départ : {choix_metier}", format_titre)
                worksheet.write("A2", f"Dimensions sélectionnées : {categories_str}", format_soustitre)
                # Définir les pondérations appliquées (selon sélection)
                pond_list = []
                if avec_savoir_faire:
                    pond_list.append(f"Savoir-faire = {poids_sf}%")
                if avec_savoir_etre:
                    pond_list.append(f"Savoir-être professionnels = {poids_se}%")
                if avec_savoirs:
                    pond_list.append(f"Savoirs = {poids_savoirs}%")

                pond_str = " / ".join(pond_list)
                worksheet.write("A3", f"Pondérations appliquées : {pond_str}", format_soustitre)

                date_export = datetime.now().strftime("%d/%m/%Y à %Hh%M")
                worksheet.write("A4", f"Date d’export : {date_export}", format_soustitre)

                # Export des données à partir de la ligne 6 (index 5)
                df.to_excel(writer, index=False, startrow=5, sheet_name=titre_feuille)

                # Ajustement automatique de la largeur des colonnes
                for i, col in enumerate(df.columns):
                    column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(i, i, column_len)

            return buffer.getvalue()

        # ⬇️ Bouton 1 : Télécharger uniquement les passerelles filtrées (score > 3)
        st.download_button(
            label="📥 Télécharger les passerelles",
            data=exporter_excel(df_filtré, "Passerelles filtrées", "passerelles_filtrees.xlsx"),
            file_name="passerelles_filtrees.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        with st.expander("📦 Télécharger toutes les passerelles (brutes)"):
            if st.button("➡️ Générer toutes les passerelles sans aucun filtre"):

                st.markdown("###\n### ⏳ Génération des passerelles brutes...")

                # Rechargement brut
                df_brut = pd.read_excel(fichier_competences, sheet_name="Macro-Compétences")
                df_brut = df_brut.dropna(subset=["Code Métier", "Intitulé", "Macro Compétence"])

                df_metiers_client = pd.read_excel(fichier_client)
                codes_client_brut = df_metiers_client["Code ROME"].dropna().unique()

                df_metiers_client = df_brut[df_brut["Code Métier"].isin(codes_client_brut)]
                df_metiers_hors_client = df_brut[~df_brut["Code Métier"].isin(codes_client_brut)]

                # Fonction de calcul avec barre de progression
                def calculer_passerelles(metiers_depart, metiers_arrivee, progress_bar=None):
                    lignes = []
                    total = len(metiers_depart["Code Métier"].unique())
                    for i, (code_depart, groupe_depart) in enumerate(metiers_depart.groupby("Code Métier")):
                        intitule_depart = groupe_depart["Intitulé"].iloc[0]
                        competences_depart = set(groupe_depart["Macro Compétence"].dropna())
                        for code_arrivee, groupe_arrivee in metiers_arrivee.groupby("Code Métier"):
                            intitule_arrivee = groupe_arrivee["Intitulé"].iloc[0]
                            competences_arrivee = set(groupe_arrivee["Macro Compétence"].dropna())
                            intersection = competences_depart & competences_arrivee
                            if intersection:
                                for comp in intersection:
                                    # Cherche catégorie dans le groupe d’arrivée (prioritaire) ou de départ
                                    if comp in groupe_arrivee["Macro Compétence"].values:
                                        cat = groupe_arrivee[groupe_arrivee["Macro Compétence"] == comp]["Catégorie"].iloc[0]
                                    else:
                                        cat = groupe_depart[groupe_depart["Macro Compétence"] == comp]["Catégorie"].iloc[0]

                                    lignes.append({
                                        "Code Métier Départ": code_depart,
                                        "Intitulé Départ": intitule_depart,
                                        "Code Métier Arrivée": code_arrivee,
                                        "Intitulé Arrivée": intitule_arrivee,
                                        "Nombre de compétences partagées": len(intersection),
                                        "Catégorie": cat,
                                        "Compétence commune": comp
                                    })

                        if progress_bar:
                            progress_bar.progress((i + 1) / total)
                    return pd.DataFrame(lignes)

                # Calcul des passerelles avec barres de progression
                st.markdown("🔄 Calcul des passerelles entrantes...")
                bar1 = st.progress(0)
                df_entrantes = calculer_passerelles(df_metiers_hors_client, df_metiers_client, progress_bar=bar1)

                st.markdown("🔄 Calcul des passerelles sortantes...")
                bar2 = st.progress(0)
                df_sortantes = calculer_passerelles(df_metiers_client, df_metiers_hors_client, progress_bar=bar2)

                st.success("✅ Calcul terminé ! Prêt à télécharger")

                # Export Excel
                buffer_brut = io.BytesIO()
                with pd.ExcelWriter(buffer_brut, engine="xlsxwriter") as writer:
                    df_entrantes.to_excel(writer, index=False, sheet_name="Passerelles entrantes")
                    df_sortantes.to_excel(writer, index=False, sheet_name="Passerelles sortantes")

                st.download_button(
                    label="📥 Télécharger le fichier complet des passerelles",
                    data=buffer_brut.getvalue(),
                    file_name="passerelles_brutes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Cliquez sur le bouton ci-dessus pour lancer le calcul complet.")

    else:
        st.warning("Aucune compétence partagée trouvée avec les métiers cibles.")
