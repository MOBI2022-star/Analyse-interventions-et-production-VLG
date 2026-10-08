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



JOURS_SEMAINE = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
MANUEL = "✍️ Saisir les jours manuellement"


HEURES_JOUR_DEFAUT = 7.0  # 7 h travaillées = 1 jour complet


# Codes qui ne constituent PAS une présence travaillée.
# CP et ANR sont donc exclus du dénominateur de la moyenne.
CODES_ABSENCE = {"CP", "ANR", "ABS", "ABSENCE", "CONGE", "CONGÉ", "RTT", "MALADIE"}

def normaliser_code_planning(val):
    if val is None:
        return ""
    return str(val).strip().upper()

def valeur_en_minutes(val):
    """Convertit une saisie du planning en minutes.
    CP/ANR et autres codes d'absence = 0 minute.
    """
    code = normaliser_code_planning(val)
    if code in CODES_ABSENCE:
        return 0
    v = code.lower().replace(",", ".").replace("h", ":")
    if v in ("", "none", "nan", "nat"):
        return 0
    if v.endswith(":"):
        v += "00"
    if ":" in v:
        try:
            h, m = v.split(":")[:2]
            return int(h) * 60 + int(m or 0)
        except Exception:
            return None
    try:
        x = float(v)
    except Exception:
        return None
    if x < 0:
        return None
    if x <= 1:
        return round(x * HEURES_JOUR_DEFAUT * 60)
    return round(x * 60)

def valeur_en_jour(val, heures_jour=HEURES_JOUR_DEFAUT):
    """Retourne uniquement la fraction de présence réellement travaillée."""
    code = normaliser_code_planning(val)
    if code in CODES_ABSENCE:
        return 0.0
    minutes = valeur_en_minutes(val)
    if not minutes:
        return 0.0
    return min(minutes / (heures_jour * 60.0), 1.0)

def compter_jours_travailles(nom_mecanicien, debut, fin, heures_jour=HEURES_JOUR_DEFAUT):
    """Compte uniquement les jours réellement travaillés.
    CP, ANR et jours non renseignés = 0.
    """
    total = 0.0
    semaines_manquantes = set()
    jour = debut
    while jour <= fin:
        _, semaine, num_jour = jour.isocalendar()
        nom_jour = JOURS_SEMAINE[num_jour - 1]
        key = f"Semaine {semaine}_{nom_mecanicien}_{nom_jour}"
        if key in st.session_state["planning_db"]:
            total += valeur_en_jour(st.session_state["planning_db"][key], heures_jour)
        else:
            semaines_manquantes.add(semaine)
        jour += datetime.timedelta(days=1)
    return round(total, 2), semaines_manquantes


def trouver_mecanicien_planning(nom_fichier, noms_planning):
    """Retourne le nom du planning le plus proche du nom trouvé dans le fichier (ou None)."""
    import difflib
    norm = lambda x: re.sub(r"\s+", " ", str(x).strip().lower())
    table = {norm(n): n for n in noms_planning}
    cible = norm(nom_fichier)
    if cible in table:
        return table[cible]
    proche = difflib.get_close_matches(cible, list(table.keys()), n=1, cutoff=0.6)
    return table[proche[0]] if proche else None


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
                # Aucun jour n'est considéré comme travaillé tant qu'il n'est pas enregistré.
                st.session_state["planning_db"][key] = "00:00"
            row[j] = st.session_state["planning_db"][key]
        records.append(row)

    df_planning = pd.DataFrame(records)

    st.info("💡 **Exemples de saisie :** `07:00` (7h = 1 jour), `06:00`, `05:00`, `06:45`... ou un **code de présence** de `0.1` à `0.7` (fraction de journée : `0.5` = demi-journée, soit 3h30). `00:00` ou `0` = absent. Pour la prime : 7h = 1 jour complet.")

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
    else:
        non_enregistre = any(
            str(row[j]) != str(st.session_state["planning_db"].get(f"{semaine_sel}_{row['Mécanicien']}_{j}"))
            for _, row in edited_df.iterrows() for j in jours
        )
        if non_enregistre:
            st.warning("⚠️ Modifications non enregistrées : cliquez sur « Enregistrer les heures de la semaine » pour qu'elles soient prises en compte (planning et prime).")

    saisies_invalides = [
        f"{row['Mécanicien']} - {j} : « {row[j]} »"
        for _, row in edited_df.iterrows() for j in jours
        if valeur_en_minutes(row[j]) is None
        and normaliser_code_planning(row[j]) not in CODES_ABSENCE
    ]
    if saisies_invalides:
        st.error("Saisies non reconnues (comptées comme 0) : " + " ; ".join(saisies_invalides))

    st.divider()
    st.subheader("📊 Récapitulatif Hebdomadaire des Heures")

    def total_heures_semaine(row):
        total_minutes = sum((valeur_en_minutes(row[j]) or 0) for j in jours)
        return f"{total_minutes // 60:02d}h{total_minutes % 60:02d}"

    def total_jours_semaine(row):
        return round(sum(valeur_en_jour(row[j]) for j in jours), 2)

    df_recap = edited_df.copy()
    df_recap["Total Semaine"] = df_recap.apply(total_heures_semaine, axis=1)
    df_recap["Jours travaillés (équiv.)"] = df_recap.apply(total_jours_semaine, axis=1)
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
        "💡 **Principe :** on prend le fichier du mois (un seul fichier), puis on va chercher dans le **planning** "
        "le nombre de **jours travaillés** de chaque mécanicien sur la période choisie.\n\n"
        "**Moyenne de vélos par jour réellement travaillé** = total des vélos ÷ présence travaillée. "
        "Cette moyenne passe dans le **même barème qu'avant** : moins de 6 vélos/jour = 0 € ; "
        "6 vélos/jour = 50 € (+26,25 € par vélo en plus) ; 10 vélos/jour et plus = 155 € max. "
        "**Prime mensuelle** = prime du jour × jours travaillés. "
        "Les journées partielles comptent en fraction de jour (7h = 1 jour ; 3h30 ou code 0.5 = 0,5 jour). Les CP et ANR sont exclus du calcul."
    )

    fichier_mois = st.file_uploader("📂 Injecter le fichier du mois (Excel / CSV)", type=["xls", "xlsx", "csv"], key="fichier_mois")

    aujourd_hui = datetime.date.today()
    debut_def = aujourd_hui.replace(day=1)
    fin_def = (debut_def + datetime.timedelta(days=32)).replace(day=1) - datetime.timedelta(days=1)

    periode = st.date_input("📆 Période du planning à prendre en compte (du ... au ...)", value=(debut_def, fin_def), format="DD/MM/YYYY")

    heures_jour = st.number_input("⏱️ Nombre d'heures correspondant à 1 jour complet", min_value=1.0, max_value=24.0, value=HEURES_JOUR_DEFAUT, step=0.5)

    if fichier_mois is None:
        st.warning("Injectez le fichier du mois pour lancer le calcul.")
    elif not (isinstance(periode, (tuple, list)) and len(periode) == 2):
        st.warning("Sélectionnez une date de début ET une date de fin.")
    else:
        debut, fin = periode
        try:
            df_m = lire_fichier(fichier_mois)
        except Exception as e:
            st.error(f"Erreur lors de la lecture du fichier : {e}")
            df_m = None

        if df_m is not None:
            if ACTOR_COL not in df_m.columns or MAT_COL not in df_m.columns:
                st.error(f"Colonnes « {ACTOR_COL} » et/ou « {MAT_COL} » introuvables dans le fichier.")
            else:
                # Filtre optionnel par date d'intervention
                choix_col = st.selectbox(
                    "🗓️ Colonne de date des interventions (optionnel, pour ne garder que la période)",
                    ["(aucune - utiliser tout le fichier)"] + list(df_m.columns)
                )
                df_f = df_m.copy()
                if not choix_col.startswith("(aucune"):
                    dates = pd.to_datetime(df_f[choix_col], dayfirst=True, errors="coerce").dt.date
                    df_f = df_f[(dates >= debut) & (dates <= fin)]

                velos = df_f.groupby(ACTOR_COL)[MAT_COL].count()
                noms_planning = [u["nom"] for u in st.session_state["users_db"].values()]

                st.subheader("🔗 Association des agents du fichier avec le planning")
                st.caption("Choisissez le mécanicien du planning correspondant à chaque agent du fichier, ou saisissez les jours à la main.")

                lignes = []
                semaines_manquantes = set()
                for acteur, nb_velos in velos.items():
                    c1, c2 = st.columns([2, 2])
                    c1.markdown(f"**{acteur}** : {int(nb_velos)} vélos")
                    options = noms_planning + [MANUEL]
                    defaut = trouver_mecanicien_planning(acteur, noms_planning)
                    idx = options.index(defaut) if defaut else len(options) - 1
                    choix = c2.selectbox("Planning", options, index=idx, key=f"map_{acteur}", label_visibility="collapsed")

                    if choix == MANUEL:
                        jours_trav = float(st.number_input(f"Jours travaillés - {acteur}", min_value=0.0, max_value=366.0, value=0.0, step=0.5, key=f"jt_{acteur}"))
                    else:
                        jours_trav, manq = compter_jours_travailles(choix, debut, fin, heures_jour)
                        semaines_manquantes |= manq

                    moyenne = nb_velos / jours_trav if jours_trav > 0 else 0.0
                    prime_jour = calculer_prime(moyenne)
                    lignes.append({
                        "Mécanicien": acteur,
                        "Total vélos": int(nb_velos),
                        "Jours travaillés": jours_trav,
                        "Moyenne vélos / jour": round(moyenne, 2),
                        "Prime / jour (€)": prime_jour,
                        "Prime mensuelle (€)": round(prime_jour * jours_trav, 2)
                    })

                if semaines_manquantes:
                    st.warning(
                        "⚠️ Ces semaines n'ont jamais été ouvertes dans le planning : "
                        + ", ".join(f"S{w}" for w in sorted(semaines_manquantes))
                        + ". Valeurs par défaut utilisées (lundi à vendredi = 1 jour). Vérifiez-les dans la section Planning."
                    )

                if lignes:
                    recap = pd.DataFrame(lignes).sort_values("Prime mensuelle (€)", ascending=False)

                    st.divider()
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Période", f"{debut.strftime('%d/%m/%Y')} → {fin.strftime('%d/%m/%Y')}")
                    m2.metric("Total vélos", int(recap["Total vélos"].sum()))
                    m3.metric("Total primes", f"{recap['Prime mensuelle (€)'].sum():.2f} €")

                    st.subheader("Prime mensuelle par mécanicien")
                    st.dataframe(
                        recap.style.format({
                            "Jours travaillés": "{:.2f}",
                            "Moyenne vélos / jour": "{:.2f}",
                            "Prime / jour (€)": "{:.2f} €",
                            "Prime mensuelle (€)": "{:.2f} €"
                        }),
                        use_container_width=True
                    )

                    st.download_button(
                        label="📥 Télécharger le rapport mensuel (Excel)",
                        data=to_excel(recap),
                        file_name=f"prime_mensuelle_{debut}_{fin}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("Aucune intervention trouvée sur la période.")
