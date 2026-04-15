#!/usr/bin/env python3
"""
Complete GLPI Data Processing & Visualization Pipeline
Generates PNG slides from JSON data for V1 and V2 sources
"""

import json
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path
from datetime import date
import unicodedata
import warnings
import os

warnings.filterwarnings('ignore')

def normalize_column_name(name: str) -> str:
    if not isinstance(name, str):
        return ''
    normalized = unicodedata.normalize('NFKD', name)
    normalized = normalized.encode('ascii', 'ignore').decode('ascii')
    return normalized.lower().strip()


# Color scheme
BLUE = '#1F4E79'
LBLUE = '#2E75B6'
LGREEN = '#4CAF50'
RED = '#C00000'
ORANGE = '#E67E22'
GRAY = '#7F7F7F'
DARK_GRAY = '#333333'

GROUP_COLORS = {
    'DC-Cloud-Platform_Team': '#2E75B6',
    'DC-Unix_Team': '#E67E22',
    'Non affecté': '#7F7F7F',
}

# Configure matplotlib
plt.rcParams.update({
    'font.family': 'DejaVu Sans',
    'font.size': 10,
    'axes.titlesize': 13,
    'axes.labelsize': 10,
    'figure.facecolor': '#F4F6F8',
    'axes.facecolor': '#FFFFFF',
    'axes.grid': True,
    'grid.alpha': 0.3,
    'axes.spines.top': False,
    'axes.spines.right': False,
})


def load_and_prepare_data(file_path, label=""):
    """Load JSON data and prepare for analysis"""
    path = Path(file_path)
    print(f"\n📂 Loading {label}: {path.name}...")
    
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    try:
        df = pd.read_json(path, encoding='utf-8')
    except ValueError:
        df = pd.read_json(path, encoding='utf-8', lines=True)
    
    print(f"   Shape: {df.shape[0]} rows × {df.shape[1]} columns")
    
    # AUDIT DE VALIDATION - Nombre total de tickets
    print(f"   📊 AUDIT: {len(df)} tickets chargés depuis le JSON")
    
    # Clean columns
    df.columns = df.columns.str.strip()
    df = df.drop(columns=[c for c in df.columns if c.strip() == ''], errors='ignore')
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    
    # Standardize column names
    rename_map = {}
    for old_col in df.columns:
        norm = normalize_column_name(old_col)
        if 'demandeur' in norm and '- ' in old_col.lower():
            rename_map[old_col] = 'Demandeur'
        elif 'groupe' in norm and 'technicien' in norm:
            rename_map[old_col] = 'Groupe'
        elif "d'ouverture" in norm or 'date d ouverture' in norm or 'date ouverture' in norm:
            rename_map[old_col] = 'Date_Ouverture'
        elif ('resolution' in norm or 'rsolution' in norm) and 'date' in norm:
            rename_map[old_col] = 'Date_Resolution'
        elif 'cloture' in norm:
            rename_map[old_col] = 'Date_Cloture'
        elif 'derniere' in norm or 'derniere modification' in norm:
            rename_map[old_col] = 'Derniere_Modif'
        elif 'priorit' in norm or 'priority' in norm:
            rename_map[old_col] = 'Priorité'
        elif norm in ['statut', 'status', 'etat']:
            rename_map[old_col] = 'Statut'
        elif (('resolution' in norm or 'rsolution' in norm) and 'depasse' in norm) or 'sla' in norm:
            rename_map[old_col] = 'SLA_Depasse'
        elif norm in ['id', 'ticket id', 'ticket_id']:
            rename_map[old_col] = 'ID'
        elif 'time to' in norm or norm == 'ttr':
            rename_map[old_col] = 'TTR'
    
    df = df.rename(columns=rename_map)
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    
    # Parse dates
    for col in ['Date_Ouverture', 'Date_Resolution', 'Date_Cloture', 'Derniere_Modif']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
    
    # Add missing columns
    if 'ID' not in df.columns:
        df['ID'] = range(len(df))
    if 'Demandeur' not in df.columns:
        df['Demandeur'] = ''
    if 'Priorité' not in df.columns:
        df['Priorité'] = 'Moyenne'
    if 'Statut' not in df.columns:
        df['Statut'] = 'Nouveau'
    if 'Groupe' not in df.columns:
        df['Groupe'] = 'Non affecté'
    if 'SLA_Depasse' not in df.columns:
        df['SLA_Depasse'] = 'Non'
    
    # AUDIT DE VALIDATION - Valeurs uniques dans Priorité
    if 'Priorité' in df.columns:
        unique_priorites = df['Priorité'].astype(str).str.strip().unique()
        print(f"   📊 AUDIT: Valeurs uniques dans 'Priorité': {list(unique_priorites)}")
    
    # Derive features
    df['Source'] = df['Demandeur'].fillna('').apply(
        lambda x: 'Automatique' if 'monito_twin' in str(x).lower() else 'Utilisateur')
    
    # SÉCURITÉ MAPPING - Transformer en minuscule et supprimer espaces
    df['Priorite_Code'] = df['Priorité'].astype(str).str.strip().str.lower().map({
        'critique': 'P1', 'haute': 'P2', 'moyenne': 'P3', 'basse': 'P4',
        'p1': 'P1', 'p2': 'P2', 'p3': 'P3', 'p4': 'P4',
        '1': 'P1', '2': 'P2', '3': 'P3', '4': 'P4'
    }).fillna('Autre')
    
    df['Groupe'] = df['Groupe'].fillna('Non affecté').astype(str).replace('', 'Non affecté')
    df['Est_Clos'] = df['Statut'].astype(str).str.lower().isin(['clos', 'résolu', 'closed', 'resolved'])
    
    df['TTR_heures'] = 0.0
    df['TTO_heures'] = 0.0
    if 'Date_Resolution' in df.columns and 'Date_Ouverture' in df.columns:
        df['TTR_heures'] = (df['Date_Resolution'] - df['Date_Ouverture']).dt.total_seconds() / 3600
        df['TTR_heures'] = df['TTR_heures'].clip(lower=0)
    if 'Derniere_Modif' in df.columns and 'Date_Ouverture' in df.columns:
        df['TTO_heures'] = (df['Derniere_Modif'] - df['Date_Ouverture']).dt.total_seconds() / 3600
        df['TTO_heures'] = df['TTO_heures'].clip(lower=0)
    
    TODAY = pd.Timestamp(date.today())
    if 'Date_Ouverture' in df.columns:
        df['Age_Jours'] = (TODAY - df['Date_Ouverture']).dt.days.clip(lower=0)
        df['Mois'] = df['Date_Ouverture'].dt.to_period('M')
    else:
        df['Age_Jours'] = 0
        df['Mois'] = pd.Period(TODAY, 'M')
    
    # AUDIT DE VALIDATION - Tickets filtrés pour le mois en cours
    current_month = df['Mois'].max() if len(df) > 0 else None
    if current_month is not None:
        df_current = df[df['Mois'] == current_month]
        print(f"   📊 AUDIT: {len(df_current)} tickets filtrés pour le mois en cours ({current_month})")
        if len(df_current) == 0:
            print("   ⚠️  ALERTE: Aucun ticket trouvé pour le mois en cours!")
    else:
        print("   ⚠️  ALERTE: Impossible de déterminer le mois en cours!")
    
    print(f"   ✓ Data prepared: {current_month if current_month else 'N/A'}")
    return df


def save_slide1_synthese(df, suffix, output_dir: Path):
    """Generate Slide 1: Daily summary with groups"""
    if len(df) == 0:
        print(f"   ⚠ Skipped slide1_courbe_groupes_{suffix}.png (empty data)")
        return
    
    nb_jours = 31
    if 'Date_Ouverture' in df.columns and df['Date_Ouverture'].notna().any():
        nb_jours = int(df['Date_Ouverture'].dt.days_in_month.max())
    
    # Created per day
    if 'Date_Ouverture' in df.columns:
        crees_jour = df.assign(Jour=df['Date_Ouverture'].dt.day).groupby('Jour').size()
    else:
        crees_jour = pd.Series(dtype=int)
    
    # Closed per day
    clot_jour = pd.Series(dtype=int)
    if 'Date_Resolution' in df.columns and df['Date_Resolution'].notna().any():
        closed_df = df[df['Date_Resolution'].notna()]
        if len(closed_df) > 0:
            clot_jour = closed_df.assign(Jour=closed_df['Date_Resolution'].dt.day).groupby('Jour').size()
    elif 'Date_Cloture' in df.columns and df['Date_Cloture'].notna().any():
        closed_df = df[df['Date_Cloture'].notna()]
        if len(closed_df) > 0:
            clot_jour = closed_df.assign(Jour=closed_df['Date_Cloture'].dt.day).groupby('Jour').size()
    
    daily = pd.DataFrame(index=range(1, nb_jours + 1))
    daily = daily.join(crees_jour.rename('Créés')).join(clot_jour.rename('Clôturés')).fillna(0)
    
    fig, axes = plt.subplots(1, 2, figsize=(16, 5.5))
    fig.patch.set_facecolor('#F4F6F8')
    month_label = str(df['Mois'].max()) if 'Mois' in df.columns and len(df) > 0 else 'Données'
    fig.suptitle(f"Synthèse Mensuelle — {month_label}", fontsize=16, fontweight='bold', color=BLUE)
    
    # Daily trend
    ax1 = axes[0]
    ax1.plot(daily.index, daily['Créés'], color=LBLUE, lw=2.5, marker='o', ms=5, label='Créés')
    ax1.plot(daily.index, daily['Clôturés'], color=LGREEN, lw=2.5, marker='s', ms=5, label='Clôturés')
    ax1.fill_between(daily.index, daily['Créés'], alpha=0.12, color=LBLUE)
    ax1.fill_between(daily.index, daily['Clôturés'], alpha=0.12, color=LGREEN)
    ax1.set_title('Évolution Journalière', fontweight='bold', color=BLUE)
    ax1.set_xlabel('Jour du mois')
    ax1.set_ylabel('Tickets')
    ax1.legend()
    ax1.set_xlim(1, nb_jours)
    
    # Groups
    ax2 = axes[1]
    if 'Groupe' in df.columns:
        gm = df['Groupe'].value_counts().sort_values()
        if len(gm) > 0:
            bars = ax2.barh(gm.index, gm.values, color=[GROUP_COLORS.get(g, GRAY) for g in gm.index])
            ax2.set_xlim(0, gm.max() * 1.2)
    ax2.set_title("Tickets par Groupe", fontweight='bold', color=BLUE)
    ax2.set_xlabel('Nombre')
    
    plt.tight_layout()
    out_path = output_dir / f'slide1_courbe_groupes_{suffix}.png'
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"   ✓ {out_path}")


def save_slide1_source(df, suffix, output_dir: Path):
    """Generate Slide 1: Source and Status"""
    if len(df) == 0:
        print(f"   ⚠ Skipped slide1_source_{suffix}.png (empty data)")
        return
    
    fig, axes = plt.subplots(1, 2, figsize=(13, 5))
    fig.patch.set_facecolor('#F4F6F8')
    fig.suptitle('Source & Statut', fontsize=14, fontweight='bold', color=BLUE)
    
    # Source pie
    ax1 = axes[0]
    if 'Source' in df.columns:
        src = df['Source'].value_counts()
        colors = [LBLUE if s == 'Automatique' else ORANGE for s in src.index]
        ax1.pie(src.values, labels=src.index, autopct='%1.1f%%', colors=colors,
                startangle=90, wedgeprops={'edgecolor': 'white', 'linewidth': 2})
    ax1.set_title('Source', fontweight='bold', color=BLUE)
    
    # Status bar
    ax2 = axes[1]
    if 'Statut' in df.columns:
        statuts = df['Statut'].value_counts()
        colors = {'Clos': LGREEN, 'Résolu': LBLUE, 'Nouveau': GRAY}
        bar_colors = [colors.get(s, GRAY) for s in statuts.index]
        ax2.bar(statuts.index, statuts.values, color=bar_colors, edgecolor='white', width=0.6)
        ax2.tick_params(axis='x', rotation=15)
    ax2.set_title('Statut', fontweight='bold', color=BLUE)
    ax2.set_ylabel('Tickets')
    
    plt.tight_layout()
    out_path = output_dir / f'slide1_source_{suffix}.png'
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"   ✓ {out_path}")


def save_slide2_jauges(df, suffix, output_dir: Path):
    """Generate Slide 2: SLA Gauges"""
    fig, axes = plt.subplots(1, 3, figsize=(15, 5.5))
    fig.patch.set_facecolor('#F4F6F8')
    fig.suptitle('Taux de Résolution SLA — P1/P2/P3', fontsize=15, fontweight='bold', color=BLUE)
    
    for idx, (ax, code, label) in enumerate(zip(axes, ['P1', 'P2', 'P3'], ['P1 — Critique', 'P2 — Haute', 'P3 — Moyenne'])):
        if 'Priorite_Code' not in df.columns:
            ax.text(0.5, 0.5, 'N/A', ha='center', va='center', fontsize=20)
            ax.axis('off')
            continue
            
        sub = df[df['Priorite_Code'] == code]
        total_tickets = len(sub)
        if total_tickets == 0:
            pct = 0
            resolved_on_time = 0
        else:
            resolved_on_time = (sub['SLA_Depasse'].astype(str) == 'Non').sum()
            pct = resolved_on_time / total_tickets * 100
        
        # DEBUG INFORMATION - Afficher dans la console
        print(f"   🔍 DEBUG SLA {code}: {resolved_on_time}/{total_tickets} ({pct:.1f}%)")
        
        is_ok = pct >= 90
        color = LGREEN if is_ok else RED
        ax.set_facecolor('#E8F5E9' if is_ok else '#FFEBEE')
        
        # Draw gauge
        th_bg = np.linspace(np.pi, 0, 300)
        ax.plot(np.cos(th_bg), np.sin(th_bg), lw=22, color='#E0E0E0')
        th_v = np.linspace(np.pi, np.pi - pct / 100 * np.pi, 300)
        ax.plot(np.cos(th_v), np.sin(th_v), lw=22, color=color)
        
        ax.text(0, -0.05, f'{pct:.1f}%', ha='center', va='center', fontsize=24, fontweight='bold', color=color)
        ax.text(0, -0.38, label, ha='center', va='center', fontsize=12, fontweight='bold', color=BLUE)
        
        # COMMENTAIRES DYNAMIQUES - Texte sous la jauge
        if total_tickets == 0:
            comment = "Aucun ticket"
        elif pct >= 95:
            comment = "Excellent"
        elif pct >= 90:
            comment = "Bon niveau"
        elif pct >= 80:
            comment = "À améliorer"
        elif pct >= 70:
            comment = "Critique"
        else:
            comment = "Urgent"
        
        ax.text(0, -0.65, f"{resolved_on_time}/{total_tickets} ({pct:.1f}%)", 
                ha='center', va='center', fontsize=10, color=DARK_GRAY)
        ax.text(0, -0.85, comment, ha='center', va='center', fontsize=9, 
                fontweight='bold', color=color)
        
        ax.set_xlim(-1.35, 1.35)
        ax.set_ylim(-0.9, 1.15)
        ax.set_aspect('equal')
        ax.axis('off')
    
    plt.tight_layout()
    out_path = output_dir / f'slide2_jauges_{suffix}.png'
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"   ✓ {out_path}")


def save_slide2_tableau(df, suffix, output_dir: Path):
    """Generate Slide 2: Group×Priority table"""
    if 'Groupe' not in df.columns or 'Priorite_Code' not in df.columns:
        print(f"   ⚠ Skipped slide2_tableau_{suffix}.png (missing columns)")
        return
    
    pivot = df.pivot_table(index='Groupe', columns='Priorite_Code', values='ID', aggfunc='count', fill_value=0)
    for col in ['P1', 'P2', 'P3']:
        if col not in pivot.columns:
            pivot[col] = 0
    pivot = pivot[['P1', 'P2', 'P3']].astype(int)
    pivot['Total'] = pivot['P1'] + pivot['P2'] + pivot['P3']
    pivot = pivot.sort_values('Total', ascending=False)
    
    fig, ax = plt.subplots(figsize=(11, max(2.5, len(pivot) * 0.8)))
    fig.patch.set_facecolor('#F4F6F8')
    ax.axis('off')
    
    reset = pivot.reset_index()
    tbl = ax.table(cellText=reset.values, colLabels=reset.columns, loc='center', cellLoc='center')
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(11)
    tbl.scale(1.3, 2.0)
    
    for j in range(len(reset.columns)):
        tbl[0, j].set_facecolor(BLUE)
        tbl[0, j].set_text_props(color='white', fontweight='bold')
    
    ax.set_title('Tickets par Groupe et Priorité', fontweight='bold', color=BLUE, pad=10)
    
    plt.tight_layout()
    out_path = output_dir / f'slide2_tableau_{suffix}.png'
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"   ✓ {out_path}")


def save_slide3_historique(df, suffix, output_dir: Path):
    """Generate Slide 3: Monthly history"""
    if 'Date_Ouverture' not in df.columns:
        print(f"   ⚠ Skipped slide3_historique_{suffix}.png (missing Date_Ouverture)")
        return
    
    monthly_open = df.groupby(df['Date_Ouverture'].dt.to_period('M')).size()
    
    if 'Date_Resolution' in df.columns:
        closed_df = df[df['Date_Resolution'].notna()]
        if len(closed_df) > 0:
            monthly_close = closed_df.groupby(closed_df['Date_Resolution'].dt.to_period('M')).size()
        else:
            monthly_close = pd.Series(dtype=int)
    elif 'Date_Cloture' in df.columns:
        closed_df = df[df['Date_Cloture'].notna()]
        if len(closed_df) > 0:
            monthly_close = closed_df.groupby(closed_df['Date_Cloture'].dt.to_period('M')).size()
        else:
            monthly_close = pd.Series(dtype=int)
    else:
        monthly_close = pd.Series(dtype=int)
    
    monthly = pd.DataFrame({'Créés': monthly_open, 'Clôturés': monthly_close}).fillna(0)
    idx_str = [str(m) for m in monthly.index]
    
    fig, ax = plt.subplots(figsize=(14, 5))
    fig.patch.set_facecolor('#F4F6F8')
    fig.suptitle(f'Historique Mensuel — {idx_str[0] if idx_str else "N/A"}', fontsize=15, fontweight='bold', color=BLUE)
    
    x = range(len(monthly))
    ax.plot(x, monthly['Créés'], color=LBLUE, lw=2.5, marker='o', ms=6, label='Créés')
    ax.plot(x, monthly['Clôturés'], color=LGREEN, lw=2.5, marker='s', ms=6, label='Clôturés')
    ax.fill_between(x, monthly['Créés'], alpha=0.12, color=LBLUE)
    ax.fill_between(x, monthly['Clôturés'], alpha=0.12, color=LGREEN)
    ax.set_xticks(list(x))
    ax.set_xticklabels(idx_str)
    ax.set_title('Évolution', fontweight='bold', color=BLUE)
    ax.set_ylabel('Tickets')
    ax.legend()
    
    plt.tight_layout()
    out_path = output_dir / f'slide3_historique_{suffix}.png'
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"   ✓ {out_path}")


def save_slide3_backlog_ttr(df, suffix, output_dir: Path):
    """Generate Slide 3: Backlog and TTR/TTO"""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
    fig.patch.set_facecolor('#F4F6F8')
    fig.suptitle('Backlog & Temps de Résolution', fontsize=13, fontweight='bold', color=BLUE)
    
    # Backlog table
    if 'Age_Jours' in df.columns:
        open_df = df[~df['Est_Clos']]
        backlog = {
            '>3j': (open_df['Age_Jours'] > 3).sum(),
            '>5j': (open_df['Age_Jours'] > 5).sum(),
            '>10j': (open_df['Age_Jours'] > 10).sum(),
            '>20j': (open_df['Age_Jours'] > 20).sum(),
        }
    else:
        backlog = {'>3j': 0, '>5j': 0, '>10j': 0, '>20j': 0}
    
    ax1.axis('off')
    bl_rows = list(backlog.items())
    tbl1 = ax1.table(cellText=[[k, str(v)] for k, v in bl_rows],
                     colLabels=['Age', 'Tickets'],
                     loc='center', cellLoc='center')
    tbl1.auto_set_font_size(False)
    tbl1.set_fontsize(11)
    tbl1.scale(1.3, 2.5)
    tbl1[0, 0].set_facecolor(BLUE)
    tbl1[0, 1].set_facecolor(BLUE)
    tbl1[0, 0].set_text_props(color='white', fontweight='bold')
    tbl1[0, 1].set_text_props(color='white', fontweight='bold')
    ax1.set_title('Backlog', fontweight='bold', color=BLUE, pad=10)
    
    # TTR/TTO table
    if 'TTR_heures' in df.columns and 'Priorite_Code' in df.columns:
        ttr_data = []
        for code in ['P1', 'P2', 'P3', 'Autre']:
            sub = df[df['Priorite_Code'] == code]
            if len(sub) > 0:
                ttr_val = sub['TTR_heures'].mean()
            else:
                ttr_val = 0
            ttr_data.append([code, f'{ttr_val:.1f}h'])
    else:
        ttr_data = [['P1', 'N/A'], ['P2', 'N/A'], ['P3', 'N/A']]
    
    ax2.axis('off')
    tbl2 = ax2.table(cellText=ttr_data,
                     colLabels=['Priorité', 'TTR Moyen'],
                     loc='center', cellLoc='center')
    tbl2.auto_set_font_size(False)
    tbl2.set_fontsize(11)
    tbl2.scale(1.3, 2.5)
    tbl2[0, 0].set_facecolor(BLUE)
    tbl2[0, 1].set_facecolor(BLUE)
    tbl2[0, 0].set_text_props(color='white', fontweight='bold')
    tbl2[0, 1].set_text_props(color='white', fontweight='bold')
    ax2.set_title('TTR par Priorité', fontweight='bold', color=BLUE, pad=10)
    
    plt.tight_layout()
    out_path = output_dir / f'slide3_backlog_ttr_{suffix}.png'
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"   ✓ {out_path}")


def main(v1_path='csvjson.json', v2_path='Synthèse DC - Incident.json', output_dir='slides'):
    """Main processing pipeline"""
    csv_source = Path(v1_path)
    json_source = Path(v2_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print("\n" + "="*80)
    print("🔄 GLPI DATA VISUALIZATION PIPELINE")
    print("="*80)
    
    sources = [
        (csv_source, 'V1', str(csv_source.name)),
        (json_source, 'V2', str(json_source.name)),
    ]
    
    for source_path, suffix, label in sources:
        print(f"\n{'─'*80}")
        print(f"📊 PROCESSING: {label} as {suffix}")
        print(f"{'─'*80}")
        
        try:
            df = load_and_prepare_data(source_path, label)
            
            print(f"\n🎨 Generating visualizations...")
            save_slide1_synthese(df, suffix, output_dir)
            save_slide1_source(df, suffix, output_dir)
            save_slide2_jauges(df, suffix, output_dir)
            save_slide2_tableau(df, suffix, output_dir)
            save_slide3_historique(df, suffix, output_dir)
            save_slide3_backlog_ttr(df, suffix, output_dir)
            
            print(f"✅ {suffix} processing complete!")
            
        except Exception as e:
            print(f"❌ ERROR processing {label}: {e}")
            import traceback
            traceback.print_exc()
    
    print(f"\n{'='*80}")
    print("📋 PNG FILES READY FOR POWERPOINT")
    print(f"{'='*80}")
    
    files = [
        ('slide1_courbe_groupes_V1.png', 'Slide 1 — Courbe + Groupes V1'),
        ('slide1_source_V1.png',         'Slide 1 — Source & Statuts V1'),
        ('slide2_jauges_V1.png',         'Slide 2 — Jauges SLA V1'),
        ('slide2_tableau_V1.png',        'Slide 2 — Tableau V1'),
        ('slide3_historique_V1.png',     'Slide 3 — Historique V1'),
        ('slide3_backlog_ttr_V1.png',    'Slide 3 — Backlog + TTR V1'),
        ('slide1_courbe_groupes_V2.png', 'Slide 1 — Courbe + Groupes V2'),
        ('slide1_source_V2.png',         'Slide 1 — Source & Statuts V2'),
        ('slide2_jauges_V2.png',         'Slide 2 — Jauges SLA V2'),
        ('slide2_tableau_V2.png',        'Slide 2 — Tableau V2'),
        ('slide3_historique_V2.png',     'Slide 3 — Historique V2'),
        ('slide3_backlog_ttr_V2.png',    'Slide 3 — Backlog + TTR V2'),
    ]
    
    for fname, desc in files:
        path = output_dir / fname
        ok = path.exists()
        size = f'{path.stat().st_size/1024:.0f} KB' if ok else 'MISSING'
        status = '✓' if ok else '✗'
        print(f"  {status}  {str(path):35s} {size:10s} — {desc}")
    
    print(f"\n{'='*80}")
    print("🚀 Next step: python generate_pptx_from_pngs.py")
    print(f"{'='*80}\n")


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Generate GLPI visualizations from JSON data.')
    parser.add_argument('--data-file-v1', default='csvjson.json', help='Path to the V1 JSON file.')
    parser.add_argument('--data-file-v2', default='Synthèse DC - Incident.json', help='Path to the V2 JSON file.')
    parser.add_argument('--output-dir', default='slides', help='Folder where PNGs will be stored.')
    args = parser.parse_args()

    main(args.data_file_v1, args.data_file_v2, args.output_dir)
