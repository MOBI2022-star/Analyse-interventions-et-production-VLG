import streamlit as st
import pandas as pd
import datetime
import io

st.set_page_config(page_title="Analyseur & Planning de Présence", layout="wide")

# --- BASE DE DONNÉES UTILISATEURS ET PLANNING EN SESSION ---
if "users_db" not in st.session_state:
    st.session_state["users_db"] = {
        "admin": {"password": "AdminPassword123!", "role": "Admin", "nom": "Administrateur Principal"},
        "mecanicien1": {"password": "User2026!", "role": "Membre", "nom": "Mécanicien 1"}
    }

if "authenticated_user" not in st.session_state:
    st.session_state["authenticated_user"] = None

# Initialisation du planning annuel (S1 à S52)
if "planning_db" not in st.session_state:
    st.session_state["planning_db"] = {}

# --- ÉCRAN DE CONNEXION ---
if st.session_state["authenticated_user"] is None:
    st.title("🔒 Connexion à l'Application")
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

# --- FONCTION DE CALCUL DE LA PRIME ---
def calculer_prime(velos):
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
            return h + m/60.0 + s/3600.0
        except:
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

# --- NAVIGATION DE L'APPLICATION ---
st.title("📊 Application de Gestion - Analyse, Primes & Planning")

menu_option = st.radio(
    "📌 Choisissez une section :", 
    ["📅 Planning & Saisie de Présence (S1 à S52)", "📈 Analyse Quotidienne (Excel/CSV)"],
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
    
    # Liste des mécaniciens enregistrés
    liste_mecaniciens = [u["nom"] for u in st.session_state["users_db"].values()]
    
    jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

    st.subheader(f"📝 Saisie pour la {semaine_sel}")

    # Création ou récupération de la table de la semaine
    records = []
    for mement in liste_mecaniciens:
        row = {"Mécanicien": mement}
        for j in jours:
            key = f"{semaine_sel}_{mement}_{j}"
            if key not in st.session_state["planning_db"]:
                st.session_state["planning_db"][key] = "07:00"  # Valeur par défaut : 7H
            row[j] = st.session_state["planning_db"][key]
        records.append(row)

    df_planning = pd.DataFrame(records)

    st.info("💡 **Exemples de saisie :** `07:00` (7h), `06:00` (6h), `05:00` (5h) ou minutes précises comme `06:45` ou `07:30`.")

    # Éditeur interactif de données
    edited_df = st.data_editor(
        df_planning,
        use_container_width=True,
        num_rows="fixed",
        key=f"editor_{semaine_sel}"
    )

    # Sauvegarde dans la session
    if st.button("💾 Enregistrer les heures de la semaine", type="primary"):
        for _, row in edited_df.iterrows():
            m_nom = row["Mécanicien"]
            for j in jours:
                key = f"{semaine_sel}_{m_nom}_{j}"
                st.session_state["planning_db"][key] = str(row[j])
        st.success(f"Données enregistrées avec succès pour la {semaine_sel} !")

    st.divider()
    st.subheader("📊 Récapitulatif Hebdomadaire des Heures")

    # Calcul du total des heures par mécanicien
    def total_heures_semaine(row):
        total_minutes = 0
        for j in jours:
            val = str(row[j]).strip()
            if ":" in val:
                try:
                    h, m = map(int, val.split(":"))
                    total_minutes += h * 60 + m
                except:
                    pass
        tot_h = total_minutes // 60
        tot_m = total_minutes % 60
        return f"{tot_h:02d}h{tot_m:02d}"

    df_recap = edited_df.copy()
    df_recap["Total Semaine"] = df_recap.apply(total_heures_semaine, axis=1)
    st.dataframe(df_recap, use_container_width=True)

    # Export Excel du planning
    excel_planning = to_excel(df_recap)
    st.download_button(
        label="📥 Exporter le planning de la semaine (Excel)",
        data=excel_planning,
        file_name=f"planning_presence_{semaine_sel.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==========================================
# SECTION 2 : ANALYSE DES INTERVENTIONS & PRIMES
# ==========================================
else:
    st.header("📈 Analyse Quotidienne des Interventions & Primes")
    uploaded_file = st.file_uploader("📂 Injecter le fichier Excel / CSV du jour", type=["xls", "xlsx", "csv"])

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"Fichier **{uploaded_file.name}** injecté avec succès ! ({len(df)} interventions enregistrées)")
            
            actor_col = "Assigné à l'acteur Libellé"
            cat_abc_col = "Identifiant du formulaire personnalisé"
            cat_model_col = "Modèle d'équipement"
            mat_col = "Matériel"
            charge_col = "Charge prévue"
            
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
                
                def color_eligibility(val):
                    if "Droit de prime" in str(val):
                        return 'background-color: #d4edda; color: #155724; font-weight: bold;'
                    else:
                        return 'background-color: #f8d7da; color: #721c24; font-weight: bold;'
                
                styled_df = agent_bikes.style.applymap(color_eligibility, subset=['Éligibilité']).format({
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
                    
                    res_abc = ct_abc.reset_index()[['Assigné à l\'acteur Libellé', 'A', '% Cat A', 'B', '% Cat B', 'C', '% Cat C', 'Total Matériel']]
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
