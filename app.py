import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="Analyseur Quotidien d'Interventions", layout="wide")

st.title("📊 Application d'Analyse Quotidienne - Interventions & Matériels")
st.markdown("Glissez-déposez votre fichier Excel/CSV du jour pour obtenir l'analyse instantanée.")

# Fonction de conversion de la durée HH:MM:SS en heures
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

# Formatage des heures en string HHhMM
def format_hours(hours):
    h = int(hours)
    m = int(round((hours - h) * 60))
    if m == 60:
        h += 1
        m = 0
    return f"{h:02d}h{m:02d}"

# Zone de dépôt de fichier
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
        
        # KPIs Globaux
        total_acteurs = df[actor_col].nunique() if actor_col in df.columns else 0
        total_interventions = len(df)
        total_materiels = df[mat_col].nunique() if mat_col in df.columns else 0
        moyenne_velos_equipe = total_interventions / total_acteurs if total_acteurs > 0 else 0
        total_charge = df['Charge_heures'].sum()
        
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Nombre de Mécaniciens", total_acteurs)
        col2.metric("Total Vélos Sortis", total_interventions)
        col3.metric("Matériels Uniques", total_materiels)
        col4.metric("Moyenne Équipe (Vélos/Mécanicien)", f"{moyenne_velos_equipe:.2f}")
        col5.metric("Charge Totale Prévue", format_hours(total_charge))
        
        st.divider()
        
        # Onglets de navigation
        tab1, tab2, tab3, tab4 = st.tabs([
            "🚴‍♂️ Vélos Sortis & Performance Équipe", 
            "🏷️ Catégories A / B / C", 
            "📌 Synthèse & Charges", 
            "🚲 Modèles (VAE / VLS)"
        ])
        
        # 1. Onglet Vélos Sortis par Mécanicien vs Moyenne de l'Équipe
        with tab1:
            st.subheader("Analyse des Vélos Sortis par Agent et Écart avec la Moyenne de l'Équipe")
            st.info(f"💡 **Moyenne globale de l'équipe** = {total_interventions} vélos au total / {total_acteurs} mécaniciens = **{moyenne_velos_equipe:.2f} vélos/mécanicien**")
            
            agent_bikes = df.groupby(actor_col).agg(
                Velos_Sortis=(mat_col, 'count')
            ).reset_index()
            
            agent_bikes['Moyenne_Equipe'] = round(moyenne_velos_equipe, 2)
            agent_bikes['Écart / Moyenne'] = (agent_bikes['Velos_Sortis'] - moyenne_velos_equipe).round(2)
            agent_bikes = agent_bikes.sort_values(by='Velos_Sortis', ascending=False)
            
            st.dataframe(agent_bikes, use_container_width=True)
            st.bar_chart(agent_bikes.set_index(actor_col)['Velos_Sortis'])
        
        # 2. Onglet Catégories A, B, C
        with tab2:
            st.subheader("Analyse des Catégories A, B, C par acteur")
            if cat_abc_col in df.columns:
                ct_abc = pd.crosstab(df[actor_col], df[cat_abc_col])
                for col in ['A', 'B', 'C']:
                    if col not in ct_abc.columns:
                        ct_abc[col] = 0
                ct_abc = ct_abc[['A', 'B', 'C']]
                ct_abc['Total Matériel'] = ct_abc.sum(axis=1)
                
                # Calcul des pourcentages
                ct_abc['% Cat A'] = (ct_abc['A'] / ct_abc['Total Matériel'] * 100).round(1).astype(str) + '%'
                ct_abc['% Cat B'] = (ct_abc['B'] / ct_abc['Total Matériel'] * 100).round(1).astype(str) + '%'
                ct_abc['% Cat C'] = (ct_abc['C'] / ct_abc['Total Matériel'] * 100).round(1).astype(str) + '%'
                
                st.dataframe(
                    ct_abc[['A', '% Cat A', 'B', '% Cat B', 'C', '% Cat C', 'Total Matériel']], 
                    use_container_width=True
                )
                
                st.markdown("#### Moyennes globales des Catégories A, B, C par acteur")
                m_a = ct_abc['A'].mean()
                m_b = ct_abc['B'].mean()
                m_c = ct_abc['C'].mean()
                
                mc1, mc2, mc3 = st.columns(3)
                mc1.metric("Moyenne Catégorie A / acteur", f"{m_a:.2f}")
                mc2.metric("Moyenne Catégorie B / acteur", f"{m_b:.2f}")
                mc3.metric("Moyenne Catégorie C / acteur", f"{m_c:.2f}")
            else:
                st.warning("Colonne de catégorie A/B/C introuvable dans le fichier.")

        # 3. Onglet Synthèse
        with tab3:
            st.subheader("Nombre de matériels et charge prévue par acteur")
            summary = df.groupby(actor_col).agg(
                Nombre_Materiel=(mat_col, 'count'),
                Total_Charge_Heures=('Charge_heures', 'sum')
            ).reset_index()
            summary['Charge Prévue'] = summary['Total_Charge_Heures'].apply(format_hours)
            summary = summary[[actor_col, 'Nombre_Materiel', 'Charge Prévue']]
            st.dataframe(summary, use_container_width=True)

        # 4. Onglet Modèles
        with tab4:
            st.subheader("Répartition des modèles d'équipement (VAE / VLS)")
            if cat_model_col in df.columns:
                ct_model = pd.crosstab(df[actor_col], df[cat_model_col])
                st.dataframe(ct_model, use_container_width=True)

    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier : {e}")
