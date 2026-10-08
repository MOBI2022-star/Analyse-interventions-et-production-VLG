import streamlit as st
import pandas as pd
import datetime
import io
import re

NOM_SITE = "Analyseur VLG"

st.set_page_config(page_title=NOM_SITE, layout="wide")

MOIS_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet",
           "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

ACTOR_COL = "Assigné à l'acteur Libellé"
CAT_ABC_COL = "Identifiant du formulaire personnalisé"
CAT_MODEL_COL = "Modèle d'équipement"
MAT_COL = "Matériel"
CHARGE_COL = "Charge prévue"

# --- BASE DE DONNÉES UTILISATEURS, PLANNING ET HISTORIQUE EN SESSION ---
if "users_db" not in st.session_state:
    st.session_state["users_db"] = {
        "admin": {"password": "AdminPassword123!", "role": "Admin", "nom": "Administrateur Principal"},
        "mecanicien1": {"password": "User2026!", "role": "Membre", "nom": "Mécanicien 1"}
    }

if "authenticated_user" not in st.session_state:
    st.session_state["authenticated_user"] = None

if "planning_db" not in st.session_state:
    st.session_state["planning_db"] = {}

# Historique des vélos sortis : { "2026-10-08": {"Mécanicien": nb_velos, ...}, ... }
if "historique_velos" not in st.session_state:
    st.session_state["historique_velos"] = {}

# --- ÉCRAN DE CONNEXION ---
if st.session_state["authenticated_user"] is None:
    st.title(f"🔒 Connexion - {NOM_SITE}")
    st.markdown("Veuillez vous identifier pour accéder au système d'analyse et de planning.")

    col_login, _ = st.columns([1, 1])
    with col_login:
        username_input = st.text_input("Identifiant / Nom d'utilisateur")
        password_input = st.text_input("Mot de passe", type="password")

        if st.button("Se connecter", type="primary"):
            users = st.session_state["users_db"]
            if username_input in users and users[username_input]["password"] == password_input:
                st.session_state["authenticated_user"] = username_input
                st.success(f"Bienvenue {users[username_input]['nom']} !")
                st.rerun()
            else:
                st.error("Identifiant ou mot de passe incorrect.")
    st.stop()

# --- GESTION DE LA DÉCONNEXION & INFOS UTILISATEUR ---
current_user = st.session_state["authenticated_user"]
user_info = st.session_state["users_db"][current_user]
is_admin = (user_info["role"] == "Admin")

st.sidebar.markdown(f"### {NOM_SITE}")
st.sidebar.markdown(f"👤 **Connecté en tant que :** {user_info['nom']}")
st.sidebar.markdown(f"🛡️ **Rôle :** `{user_info['role']}`")

if st.sidebar.button("Déconnexion"):
    st.session_state["authenticated_user"] = None
    st.rerun()

# --- ESPACE D'ADMINISTRATION EXCLUSIF ---
if is_admin:
    st.sidebar.divider()
    st.sidebar.subheader("⚙️ Espace Administrateur")

    with st.sidebar.expander("👤 Gérer les Membres (Admin uniquement)"):
        st.markdown("##### ➕ Ajouter un membre")
        new_username = st.text_input("Nouvel Identifiant", key="new_user")
        new_name = st.text_input("Nom de l'agent", key="new_name")
        new_password = st.text_input("Mot de passe", type="password", key="new_pass")
        new_role = st.selectbox("Rôle", ["Membre", "Admin"], key="new_role")

        if st.button("Créer le membre"):
            if new_username and new_password:
                if new_username in st.session_state["users_db"]:
                    st.warning("Cet identifiant existe déjà.")
                else:
                    st.session_state["users_db"][new_username] = {
                        "password": new_password,
                        "role": new_role,
                        "nom": new_name if new_name else new_username
                    }
                    st.success(f"Membre '{new_username}' ajouté avec succès !")
                    st.rerun()
            else:
                st.error("Remplissez l'identifiant et le mot de passe.")

        st.divider()
        st.markdown("##### 🗑️ Supprimer un membre")
        existing_users = [u for u in st.session_state["users_db"].keys() if u != current_user]
        if existing_users:
            user_to_delete = st.selectbox("Sélectionner un membre à supprimer", existing_users)
            if st.button("Supprimer ce membre"):
                del st.session_state["users_db"][user_to_delete]
                st.success(f"Le membre '{user_to_delete}' a été supprimé.")
                st.rerun()
        else:
            st.info("Aucun autre membre à supprimer.")


# --- FONCTIONS ---
def calculer_prime(velos):
    """Même barème qu'avant : <6 vélos = 0 € ; 6 vélos = 50 € ; +26,25 € par vélo ; 10 vélos et plus = 155 €."""
    if velos < 6.0:
        return 0.0
    elif velos >= 10.0:
        return 155.0
    else:
        return round(50.0 + (velos - 6.0) * 26.25, 2)


def parse_time_to_hours(val):
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip()
    parts = val_str.split(':')
    if len(parts) == 3:
        try:
            h, m, s = map(float, parts)
            return h + m / 60.0 + s / 3600.0
        except Exception:
            return 0.0
    return 0.0


def format_hours(hours):
    h = int(hours)
    m = int(round((hours - h) * 60))
    if m == 60:
        h += 1
        m = 0
    return f"{h:02d}h{m:02d}"


def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Analyse')
    return output.getvalue()


def to_excel_multi(feuilles):
    """feuilles : dict {nom_feuille: DataFrame}"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for nom, d in feuilles.items():
            d.to_excel(writer, index=False, sheet_name=nom[:31])
    return output.getvalue()


def lire_fichier(f):
    if f.name.lower().endswith('.csv'):
        return pd.read_csv(f)
    return pd.read_excel(f)


def compter_velos_par_acteur(df):
    """Retourne {acteur: nombre de vélos sortis} pour un fichier journalier."""
    if ACTOR_COL not in df.columns or MAT_COL not in df.columns:
        return None
    return df.groupby(ACTOR_COL)[MAT_COL].count().to_dict()


def date_depuis_nom(nom):
    """Cherche une date AAAA-MM-JJ ou JJ-MM-AAAA (séparateurs - _ .) dans le nom du fichier."""
    m = re.search(r"(\d{4})[-_.](\d{2})[-_.](\d{2})", nom)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    m = re.search(r"(\d{2})[-_.](\d{2})[-_.](\d{4})", nom)
    if m:
        try:
            return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    return datetime.date.today()


def style_map(styler, func, subset):
    if hasattr(styler, "map"):
        return styler.map(func, subset=subset)
    return styler.applymap(func, subset=subset)


def color_eligibility(val):
    if "Droit de prime" in str(val):
        return 'background-color: #d4edda; color: #155724; font-weight: bold;'
    return 'background-color: #f8d7da; color: #721c24; font-weight: bold;'


# --- NAVIGATION DE L'APPLICATION ---
st.title(f"📊 {NOM_SITE} - Analyse, Primes & Planning")

menu_option = st.radio(
    "📌 Choisissez une section :",
    [
        "📅 Planning & Saisie de Présence (S1 à S52)",
        "📈 Analyse Quotidienne (Excel/CSV)",
        "💶 Prime Mensuelle"
    ],
    horizontal=True
)

st.divider()

# ==========================================
# SECTION 1 : PLANNING & SAISIE DES HEURES
# ==========================================
if menu_option == "📅 Planning & Saisie de Présence (S1 à S52)":
    st.header("📅 Saisie et Suivi de Présence Annuel")
    st.markdown("Sélectionnez la semaine (de la **Semaine 1** à la **Semaine 52**) et saisissez les heures travaillées par mécanicien.")

    col_semaine, col_agent = st.columns([1, 2])
    with col_semaine:
        semaine_sel = st.selectbox("📅 Semaine :", [f"Semaine {i}" for i in range(1, 53)])

    liste_mecaniciens = [u["nom"] for u in st.session_state["users_db"].values()]
    jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

    st.subheader(f"📝 Saisie pour la {semaine_sel}")

    records = []
    for mecanicien in liste_mecaniciens:
        row = {"Mécanicien": mecanicien}
        for j in jours:
            key = f"{semaine_sel}_{mecanicien}_{j}"
            if key not in st.session_state["planning_db"]:
                st.session_state["planning_db"][key] = "07:00"
            row[j] = st.session_state["planning_db"][key]
        records.append(row)

    df_planning = pd.DataFrame(records)

    st.info("💡 **Exemples de saisie :** `07:00` (7h), `06:00` (6h), `05:00` (5h) ou minutes précises comme `06:45` ou `07:30`.")

    edited_df = st.data_editor(
        df_planning,
        use_container_width=True,
        num_rows="fixed",
        key=f"editor_{semaine_sel}"
    )

    if st.button("💾 Enregistrer les heures de la semaine", type="primary"):
        for _, row in edited_df.iterrows():
            m_nom = row["Mécanicien"]
            for j in jours:
                key = f"{semaine_sel}_{m_nom}_{j}"
                st.session_state["planning_db"][key] = str(row[j])
        st.success(f"Données enregistrées avec succès pour la {semaine_sel} !")

    st.divider()
    st.subheader("📊 Récapitulatif Hebdomadaire des Heures")

    def total_heures_semaine(row):
        total_minutes = 0
        for j in jours:
            val = str(row[j]).strip()
            if ":" in val:
                try:
                    h, m = map(int, val.split(":"))
                    total_minutes += h * 60 + m
                except Exception:
                    pass
        tot_h = total_minutes // 60
        tot_m = total_minutes % 60
        return f"{tot_h:02d}h{tot_m:02d}"

    df_recap = edited_df.copy()
    df_recap["Total Semaine"] = df_recap.apply(total_heures_semaine, axis=1)
    st.dataframe(df_recap, use_container_width=True)

    excel_planning = to_excel(df_recap)
    st.download_button(
        label="📥 Exporter le planning de la semaine (Excel)",
        data=excel_planning,
        file_name=f"planning_presence_{semaine_sel.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==========================================
# SECTION 2 : ANALYSE DES INTERVENTIONS & PRIMES (JOURNALIER)
# ==========================================
elif menu_option == "📈 Analyse Quotidienne (Excel/CSV)":
    st.header("📈 Analyse Quotidienne des Interventions & Primes")
    uploaded_file = st.file_uploader("📂 Injecter le fichier Excel / CSV du jour", type=["xls", "xlsx", "csv"])

    if uploaded_file is not None:
        try:
            df = lire_fichier(uploaded_file)

            st.success(f"Fichier **{uploaded_file.name}** injecté avec succès ! ({len(df)} interventions enregistrées)")

            actor_col, cat_abc_col, cat_model_col = ACTOR_COL, CAT_ABC_COL, CAT_MODEL_COL
            mat_col, charge_col = MAT_COL, CHARGE_COL

            df['Charge_heures'] = df[charge_col].apply(parse_time_to_hours) if charge_col in df.columns else 0.0

            total_acteurs = df[actor_col].nunique() if actor_col in df.columns else 0
            total_interventions = len(df)
            total_materiels = df[mat_col].nunique() if mat_col in df.columns else 0
            moyenne_velos_equipe = total_interventions / total_acteurs if total_acteurs > 0 else 0
            total_charge = df['Charge_heures'].sum()

            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Nombre de Mécaniciens", total_acteurs)
            col2.metric("Total Vélos Sortis", total_interventions)
            col3.metric("Matériels Uniques", total_materiels)
            col4.metric("Moyenne Équipe", f"{moyenne_velos_equipe:.2f} vélos")
            col5.metric("Charge Totale Prévue", format_hours(total_charge))

            st.divider()

            tab1, tab2, tab3, tab4 = st.tabs([
                "🚴‍♂️ Vélos Sortis & Primes",
                "🏷️ Catégories A / B / C",
                "📌 Synthèse & Charges",
                "🚲 Modèles (VAE / VLS)"
            ])

            with tab1:
                st.subheader("Analyse des Vélos Sortis, Statut d'Éligibilité et Primes")
                st.info("💡 **Condition de prime :** À partir de **6 vélos sortis** (50,00 €) jusqu'à **10 vélos** (155,00 € max).")

                agent_bikes = df.groupby(actor_col).agg(
                    Velos_Sortis=(mat_col, 'count')
                ).reset_index()

                agent_bikes['Éligibilité'] = agent_bikes['Velos_Sortis'].apply(
                    lambda x: "🟢 Droit de prime" if x >= 6.0 else "🔴 Pas de prime"
                )
                agent_bikes['Montant Prime (€)'] = agent_bikes['Velos_Sortis'].apply(calculer_prime)
                agent_bikes = agent_bikes.sort_values(by='Velos_Sortis', ascending=False)

                styled_df = style_map(agent_bikes.style, color_eligibility, ['Éligibilité']).format({
                    'Montant Prime (€)': '{:.2f} €'
                })
                st.dataframe(styled_df, use_container_width=True)

                excel_data = to_excel(agent_bikes)
                st.download_button(
                    label="📥 Télécharger ce rapport avec Primes (Excel)",
                    data=excel_data,
                    file_name=f"rapport_primes_mecaniciens_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.bar_chart(agent_bikes.set_index(actor_col)['Velos_Sortis'])

                st.divider()
                st.markdown("##### 🗓️ Ajouter cette journée au calcul de la prime mensuelle")
                date_jour = st.date_input(
                    "Date de ce fichier",
                    value=date_depuis_nom(uploaded_file.name),
                    key="date_fichier_jour"
                )
                if st.button("➕ Ajouter à l'historique du mois"):
                    st.session_state["historique_velos"][date_jour.isoformat()] = compter_velos_par_acteur(df)
                    st.success(f"Journée du {date_jour.strftime('%d/%m/%Y')} ajoutée à l'historique.")

            with tab2:
                st.subheader("Analyse des Catégories A, B, C par acteur")
                if cat_abc_col in df.columns:
                    ct_abc = pd.crosstab(df[actor_col], df[cat_abc_col])
                    for col in ['A', 'B', 'C']:
                        if col not in ct_abc.columns:
                            ct_abc[col] = 0
                    ct_abc = ct_abc[['A', 'B', 'C']]
                    ct_abc['Total Matériel'] = ct_abc.sum(axis=1)

                    ct_abc['% Cat A'] = (ct_abc['A'] / ct_abc['Total Matériel'] * 100).round(1).astype(str) + '%'
                    ct_abc['% Cat B'] = (ct_abc['B'] / ct_abc['Total Matériel'] * 100).round(1).astype(str) + '%'
                    ct_abc['% Cat C'] = (ct_abc['C'] / ct_abc['Total Matériel'] * 100).round(1).astype(str) + '%'

                    res_abc = ct_abc.reset_index()[[actor_col, 'A', '% Cat A', 'B', '% Cat B', 'C', '% Cat C', 'Total Matériel']]
                    st.dataframe(res_abc, use_container_width=True)
                else:
                    st.warning("Colonne de catégorie A/B/C introuvable dans le fichier.")

            with tab3:
                st.subheader("Nombre de matériels et charge prévue par acteur")
                summary = df.groupby(actor_col).agg(
                    Nombre_Materiel=(mat_col, 'count'),
                    Total_Charge_Heures=('Charge_heures', 'sum')
                ).reset_index()
                summary['Charge Prévue'] = summary['Total_Charge_Heures'].apply(format_hours)
                summary = summary[[actor_col, 'Nombre_Materiel', 'Charge Prévue']]
                st.dataframe(summary, use_container_width=True)

            with tab4:
                st.subheader("Répartition des modèles d'équipement (VAE / VLS)")
                if cat_model_col in df.columns:
                    ct_model = pd.crosstab(df[actor_col], df[cat_model_col]).reset_index()
                    st.dataframe(ct_model, use_container_width=True)

        except Exception as e:
            st.error(f"Erreur lors du traitement du fichier : {e}")

# ==========================================
# SECTION 3 : PRIME MENSUELLE
# ==========================================
else:
    st.header("💶 Calcul de la Prime Mensuelle")
    st.info(
        "💡 **Mêmes critères que la prime journalière :** chaque jour, un mécanicien touche "
        "**0 €** s'il sort moins de 6 vélos, **50,00 €** à 6 vélos (+26,25 € par vélo supplémentaire), "
        "et **155,00 € maximum** à partir de 10 vélos. La prime du mois est la **somme des primes de chaque journée**."
    )

    fichiers = st.file_uploader(
        "📂 Injecter les fichiers journaliers du mois (un fichier par jour, plusieurs possibles)",
        type=["xls", "xlsx", "csv"],
        accept_multiple_files=True
    )

    # Données issues des fichiers injectés ici (la date est déduite du nom, modifiable)
    jours_importes = {}
    if fichiers:
        st.markdown("##### 🗓️ Date de chaque fichier")
        for f in fichiers:
            c1, c2 = st.columns([2, 1])
            c1.write(f"📄 {f.name}")
            d = c2.date_input("Date", value=date_depuis_nom(f.name), key=f"date_{f.name}", label_visibility="collapsed")
            try:
                comptes = compter_velos_par_acteur(lire_fichier(f))
                if comptes is None:
                    st.warning(f"Colonnes « {ACTOR_COL} » / « {MAT_COL} » introuvables dans {f.name}.")
                else:
                    jours_importes[d.isoformat()] = comptes
            except Exception as e:
                st.error(f"Erreur sur {f.name} : {e}")

    # Historique (journées ajoutées depuis l'analyse quotidienne) + fichiers injectés (prioritaires si même date)
    donnees = dict(st.session_state["historique_velos"])
    donnees.update(jours_importes)

    if not donnees:
        st.warning("Aucune journée disponible. Injectez des fichiers ci-dessus ou ajoutez des journées depuis l'« Analyse Quotidienne ».")
    else:
        lignes = []
        for date_iso, comptes in donnees.items():
            for acteur, nb in comptes.items():
                lignes.append({
                    "Date": datetime.date.fromisoformat(date_iso),
                    "Mécanicien": acteur,
                    "Velos_Sortis": int(nb),
                    "Prime (€)": calculer_prime(nb)
                })
        df_jours = pd.DataFrame(lignes)
        df_jours["Mois"] = df_jours["Date"].apply(lambda d: f"{d.year}-{d.month:02d}")

        mois_dispo = sorted(df_jours["Mois"].unique(), reverse=True)

        def label_mois(m):
            a, mm = m.split("-")
            return f"{MOIS_FR[int(mm) - 1]} {a}"

        mois_sel = st.selectbox("📆 Mois à calculer :", mois_dispo, format_func=label_mois)
        df_mois = df_jours[df_jours["Mois"] == mois_sel].copy()

        recap = df_mois.groupby("Mécanicien").agg(
            Jours_Comptabilisés=("Date", "nunique"),
            Total_Velos=("Velos_Sortis", "sum"),
            Jours_Avec_Prime=("Prime (€)", lambda s: int((s > 0).sum())),
            Prime_Mensuelle=("Prime (€)", "sum")
        ).reset_index().sort_values("Prime_Mensuelle", ascending=False)

        recap.columns = ["Mécanicien", "Jours comptabilisés", "Total vélos sortis",
                         "Jours avec prime", "Prime mensuelle (€)"]

        c1, c2, c3 = st.columns(3)
        c1.metric("Jours comptabilisés", df_mois["Date"].nunique())
        c2.metric("Total vélos sortis", int(df_mois["Velos_Sortis"].sum()))
        c3.metric("Total primes du mois", f"{recap['Prime mensuelle (€)'].sum():.2f} €")

        st.subheader(f"Prime mensuelle par mécanicien - {label_mois(mois_sel)}")
        st.dataframe(
            recap.style.format({"Prime mensuelle (€)": "{:.2f} €"}),
            use_container_width=True
        )

        st.subheader("Détail jour par jour")
        pivot = df_mois.pivot_table(
            index="Mécanicien", columns="Date", values="Prime (€)", aggfunc="sum", fill_value=0.0
        )
        pivot.columns = [d.strftime("%d/%m") for d in pivot.columns]
        st.dataframe(pivot.style.format("{:.2f} €"), use_container_width=True)

        detail = df_mois[["Date", "Mécanicien", "Velos_Sortis", "Prime (€)"]].sort_values(["Date", "Mécanicien"])
        excel_mois = to_excel_multi({
            "Prime mensuelle": recap,
            "Détail journalier": detail
        })
        st.download_button(
            label="📥 Télécharger le rapport mensuel (Excel)",
            data=excel_mois,
            file_name=f"prime_mensuelle_{mois_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if st.session_state["historique_velos"]:
            with st.expander("🗂️ Historique enregistré"):
                st.write(sorted(st.session_state["historique_velos"].keys()))
                if st.button("🗑️ Vider l'historique"):
                    st.session_state["historique_velos"] = {}
                    st.rerun()
