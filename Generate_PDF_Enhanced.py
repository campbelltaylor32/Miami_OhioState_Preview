#!/usr/bin/env python3
"""
Miami vs Ohio State - CFP Comprehensive Matchup Analysis
PDF Report with Percentile-Based Comparisons
"""

import pandas as pd
import numpy as np
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, 
    PageBreak, HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from datetime import datetime
import os

# ============================================================================
# CONFIGURATION
# ============================================================================
TEAM1 = "Miami"
TEAM2 = "Ohio State"
CSV_DIR = "."

# ============================================================================
# COLORS
# ============================================================================
MIAMI_ORANGE = colors.HexColor("#F47321")
MIAMI_GREEN = colors.HexColor("#005030")
OSU_SCARLET = colors.HexColor("#BB0000")
OSU_GRAY = colors.HexColor("#666666")
HEADER_BG = colors.HexColor("#1a1a2e")
LIGHT_GRAY = colors.HexColor("#f5f5f5")
DARK_TEXT = colors.HexColor("#333333")
GREEN_GOOD = colors.HexColor("#28a745")
RED_BAD = colors.HexColor("#dc3545")
YELLOW_MID = colors.HexColor("#ffc107")

# ============================================================================
# DATA LOADING
# ============================================================================

def load_csv_safe(filename):
    filepath = os.path.join(CSV_DIR, filename)
    try:
        return pd.read_csv(filepath)
    except FileNotFoundError:
        print(f"Warning: {filename} not found")
        return pd.DataFrame()
    except Exception as e:
        print(f"Error loading {filename}: {e}")
        return pd.DataFrame()

def load_all_data():
    data = {
        'comparison': load_csv_safe('Miami_vs_OhioState_Comparison.csv'),
        'vulnerabilities': load_csv_safe('Miami_vs_OhioState_Vulnerabilities.csv'),
        'strengths': load_csv_safe('Miami_vs_OhioState_Strengths.csv'),
        'edges': load_csv_safe('Miami_vs_OhioState_Edges.csv'),
        'trends': load_csv_safe('Miami_vs_OhioState_Trends.csv'),
        'game_script': load_csv_safe('Miami_vs_OhioState_GameScript.csv'),
        'quarters': load_csv_safe('Miami_vs_OhioState_Quarter_Analysis.csv'),
        'situations': load_csv_safe('Miami_vs_OhioState_Situation_Analysis.csv'),
        'field_position': load_csv_safe('Miami_vs_OhioState_Field_Position.csv'),
        'drives': load_csv_safe('Miami_vs_OhioState_Drive_Efficiency.csv'),
        'pressure': load_csv_safe('Miami_vs_OhioState_Pressure_Analysis.csv'),
        'late_close': load_csv_safe('Miami_vs_OhioState_Late_Close.csv'),
        'tempo': load_csv_safe('Miami_vs_OhioState_Tempo.csv'),
        # Line of Scrimmage Analysis
        'run_progression_quarters': load_csv_safe('Miami_vs_OhioState_Run_Progression_Quarters.csv'),
        'run_progression_trend': load_csv_safe('Miami_vs_OhioState_Run_Progression_Trend.csv'),
        'obvious_run': load_csv_safe('Miami_vs_OhioState_Obvious_Run.csv'),
        'run_defense_quarters': load_csv_safe('Miami_vs_OhioState_Run_Defense_Quarters.csv'),
        'run_defense_trend': load_csv_safe('Miami_vs_OhioState_Run_Defense_Trend.csv'),
        'run_defense_obvious': load_csv_safe('Miami_vs_OhioState_Run_Defense_Obvious.csv'),
        'protection_obvious': load_csv_safe('Miami_vs_OhioState_Protection_Obvious_Pass.csv'),
        'pass_rush_obvious': load_csv_safe('Miami_vs_OhioState_Pass_Rush_Obvious.csv'),
        'los_summary': load_csv_safe('Miami_vs_OhioState_LOS_Summary.csv'),
    }
    return data

# ============================================================================
# PERCENTILE UTILITIES
# ============================================================================

def is_defense_metric(metric_name):
    metric_lower = metric_name.lower()
    defense_keywords = ['defense', 'allowed', 'against', 'def_']
    return any(kw in metric_lower for kw in defense_keywords)

def is_defense_stuff_rate(metric_name):
    """Defense Stuff Rate is special - higher is better (unlike other defense metrics)"""
    return 'defense_stuff_rate' in metric_name.lower() or 'defense stuff rate' in metric_name.lower()

def is_offense_stuff_rate(metric_name):
    """Offense Stuff Rate - lower is better (getting stuffed is bad)"""
    return 'offense_stuff_rate' in metric_name.lower() or 'offense stuff rate' in metric_name.lower()

def is_int_rate_offense(metric_name):
    """Int Rate Offense - lower is better (throwing fewer interceptions is good)"""
    return 'int_rate_offense' in metric_name.lower() or 'int rate offense' in metric_name.lower()

def convert_to_quality_percentile(raw_pctl, metric_name):
    if pd.isna(raw_pctl):
        return np.nan
    # Defense Stuff Rate: higher is better (stuffing opponents is good) - don't invert
    if is_defense_stuff_rate(metric_name):
        return raw_pctl
    # Offense Stuff Rate: lower is better (getting stuffed is bad) - invert
    if is_offense_stuff_rate(metric_name):
        return 100 - raw_pctl
    # Int Rate Offense: lower is better (fewer interceptions) - invert
    if is_int_rate_offense(metric_name):
        return 100 - raw_pctl
    # Other defense metrics: lower is better - invert
    if is_defense_metric(metric_name):
        return 100 - raw_pctl
    # Other offense metrics: higher is better - don't invert
    return raw_pctl

def get_percentile_color(pctl):
    if pd.isna(pctl):
        return colors.grey
    if pctl >= 85:
        return GREEN_GOOD
    elif pctl >= 65:
        return colors.HexColor("#5cb85c")
    elif pctl >= 40:
        return YELLOW_MID
    elif pctl >= 20:
        return colors.HexColor("#f0ad4e")
    else:
        return RED_BAD

def format_percentile(pctl, show_quality=True):
    if pd.isna(pctl):
        return "N/A"
    pctl = round(pctl)
    if pctl >= 90:
        return f"{pctl}th ★"
    elif pctl <= 20:
        return f"{pctl}th ■"
    else:
        return f"{pctl}th"

def get_edge_from_percentiles(team1_pctl, team2_pctl):
    if pd.isna(team1_pctl) or pd.isna(team2_pctl):
        return "N/A", colors.grey
    diff = team1_pctl - team2_pctl
    if diff >= 15:
        return "MIAMI", MIAMI_GREEN
    elif diff <= -15:
        return "OSU", OSU_SCARLET
    elif diff >= 5:
        return "MIA+", MIAMI_GREEN
    elif diff <= -5:
        return "OSU+", OSU_SCARLET
    else:
        return "EVEN", OSU_GRAY

def clean_metric_name(metric):
    return (metric
            .replace('_season_avg', '')
            .replace('_', ' ')
            .replace('EPA per', 'EPA/')
            .replace('Success Rate', 'Success %')
            .title())

# ============================================================================
# STYLES
# ============================================================================

def create_styles():
    styles = getSampleStyleSheet()
    
    styles.add(ParagraphStyle(
        name='MainTitle', fontSize=26, textColor=colors.white,
        alignment=TA_CENTER, spaceAfter=6, fontName='Helvetica-Bold'
    ))
    
    styles.add(ParagraphStyle(
        name='SubTitle', fontSize=12, textColor=colors.HexColor("#cccccc"),
        alignment=TA_CENTER, spaceAfter=15
    ))
    
    styles.add(ParagraphStyle(
        name='SectionHeader', fontSize=13, textColor=HEADER_BG,
        spaceBefore=12, spaceAfter=6, fontName='Helvetica-Bold'
    ))
    
    styles.add(ParagraphStyle(
        name='SubSection', fontSize=10, textColor=DARK_TEXT,
        spaceBefore=8, spaceAfter=4, fontName='Helvetica-Bold'
    ))
    
    styles.add(ParagraphStyle(
        name='TeamMiami', fontSize=10, textColor=MIAMI_GREEN,
        fontName='Helvetica-Bold', spaceBefore=6, spaceAfter=3
    ))
    
    styles.add(ParagraphStyle(
        name='TeamOSU', fontSize=10, textColor=OSU_SCARLET,
        fontName='Helvetica-Bold', spaceBefore=6, spaceAfter=3
    ))
    
    styles.add(ParagraphStyle(
        name='Body', fontSize=9, textColor=DARK_TEXT, spaceAfter=5, leading=11
    ))
    
    styles.add(ParagraphStyle(
        name='Note', fontSize=7, textColor=OSU_GRAY, fontStyle='italic', spaceAfter=6
    ))
    
    return styles

# ============================================================================
# HELPER FUNCTION FOR TABLES
# ============================================================================

def create_comparison_table(table_data, col_widths, header_color=HEADER_BG):
    """Create a styled comparison table"""
    table = Table(table_data, colWidths=col_widths)
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), header_color),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]
    
    for i in range(1, len(table_data)):
        if i % 2 == 0:
            style_commands.append(('BACKGROUND', (0, i), (-1, i), LIGHT_GRAY))
    
    table.setStyle(TableStyle(style_commands))
    return table

# ============================================================================
# REPORT SECTIONS
# ============================================================================

def create_title_block(styles, data):
    elements = []
    
    # Add some top spacing
    elements.append(Spacer(1, 20))
    
    title_data = [
        [Paragraph("COLLEGE FOOTBALL PLAYOFF", styles['SubTitle'])],
        [Spacer(1, 8)],
        [Paragraph(f"{TEAM1.upper()} vs {TEAM2.upper()}", styles['MainTitle'])],
        [Spacer(1, 8)],
        [Paragraph("Comprehensive Percentile-Based Matchup Analysis", styles['SubTitle'])],
        [Spacer(1, 4)],
        [Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y')}", styles['SubTitle'])]
    ]
    
    title_table = Table(title_data, colWidths=[7.5*inch])
    title_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), HEADER_BG),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, 0), 15),
        ('BOTTOMPADDING', (0, -1), (-1, -1), 15),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
    ]))
    
    elements.append(title_table)
    elements.append(Spacer(1, 20))
    
    legend_text = "★ = Elite (90th+ percentile) | ■ = Vulnerability (20th or below) | All percentiles are QUALITY-ADJUSTED (higher = better)"
    elements.append(Paragraph(legend_text, styles['Note']))
    elements.append(Spacer(1, 15))
    
    return elements

def create_full_comparison(styles, data):
    """Comprehensive head-to-head comparison table"""
    elements = []
    
    elements.append(Paragraph("FULL STATISTICAL COMPARISON", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    comp_df = data['comparison']
    if comp_df.empty:
        elements.append(Paragraph("Comparison data not available.", styles['Body']))
        return elements
    
    table_data = [["METRIC", "MIAMI", "MIA %", "OSU", "OSU %", "EDGE"]]
    
    for _, row in comp_df.iterrows():
        metric = clean_metric_name(row['Metric'])
        miami_val = row.get('Miami', 0)
        osu_val = row.get('Ohio_State', 0)
        miami_raw_pctl = row.get('Miami_Pctl', 50)
        osu_raw_pctl = row.get('OSU_Pctl', 50)
        
        miami_qual_pctl = convert_to_quality_percentile(miami_raw_pctl, row['Metric'])
        osu_qual_pctl = convert_to_quality_percentile(osu_raw_pctl, row['Metric'])
        
        edge_text, _ = get_edge_from_percentiles(miami_qual_pctl, osu_qual_pctl)
        
        table_data.append([
            metric[:32],
            f"{miami_val:.3f}" if pd.notna(miami_val) else "-",
            format_percentile(miami_qual_pctl),
            f"{osu_val:.3f}" if pd.notna(osu_val) else "-",
            format_percentile(osu_qual_pctl),
            edge_text
        ])
    
    table = Table(table_data, colWidths=[2.2*inch, 0.75*inch, 0.8*inch, 0.75*inch, 0.8*inch, 0.7*inch])
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 7),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]
    
    for i in range(1, len(table_data)):
        row = comp_df.iloc[i-1]
        miami_qual = convert_to_quality_percentile(row.get('Miami_Pctl', 50), row['Metric'])
        osu_qual = convert_to_quality_percentile(row.get('OSU_Pctl', 50), row['Metric'])
        
        style_commands.append(('TEXTCOLOR', (2, i), (2, i), get_percentile_color(miami_qual)))
        style_commands.append(('TEXTCOLOR', (4, i), (4, i), get_percentile_color(osu_qual)))
        
        edge = table_data[i][5]
        if "MIA" in edge:
            style_commands.append(('TEXTCOLOR', (5, i), (5, i), MIAMI_GREEN))
            style_commands.append(('FONTNAME', (5, i), (5, i), 'Helvetica-Bold'))
        elif "OSU" in edge:
            style_commands.append(('TEXTCOLOR', (5, i), (5, i), OSU_SCARLET))
            style_commands.append(('FONTNAME', (5, i), (5, i), 'Helvetica-Bold'))
        
        if i % 2 == 0:
            style_commands.append(('BACKGROUND', (0, i), (-1, i), LIGHT_GRAY))
    
    table.setStyle(TableStyle(style_commands))
    elements.append(table)
    
    return elements

def create_matchup_matrix(styles, data):
    """Offense vs Defense matchup analysis"""
    elements = []
    
    elements.append(Paragraph("HEAD-TO-HEAD MATCHUP MATRIX", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 5))
    
    comp_df = data['comparison']
    if comp_df.empty:
        return elements
    
    # Miami Offense vs OSU Defense
    elements.append(Paragraph("Miami Offense vs Ohio State Defense", styles['TeamMiami']))
    
    offense_metrics = ['Offense_EPA_per_Play', 'Offense_EPA_per_Rush', 'Offense_EPA_per_Pass',
                       'Offense_Success_Rate', 'Offense_3rd_Down_Success', 'Offense_Scoring_Drive_Pct',
                       'Offense_Explosive_Rate', 'Offense_RedZone_TD_Rate']
    defense_metrics = ['Defense_EPA_per_Play', 'Defense_EPA_per_Rush', 'Defense_EPA_per_Pass',
                       'Defense_Success_Rate', 'Defense_3rd_Down_Success', 'Defense_Scoring_Drive_Pct',
                       'Defense_Explosive_Rate', 'Defense_RedZone_TD_Rate']
    
    table_data = [["MATCHUP", "MIA OFF %", "OSU DEF %", "GAP", "PROJ"]]
    
    for off_m, def_m in zip(offense_metrics, defense_metrics):
        off_key = f"{off_m}_season_avg"
        def_key = f"{def_m}_season_avg"
        
        miami_off_row = comp_df[comp_df['Metric'] == off_key]
        osu_def_row = comp_df[comp_df['Metric'] == def_key]
        
        if miami_off_row.empty or osu_def_row.empty:
            continue
        
        miami_off_pctl = miami_off_row['Miami_Pctl'].values[0]
        osu_def_raw = osu_def_row['OSU_Pctl'].values[0]
        osu_def_qual = convert_to_quality_percentile(osu_def_raw, def_key)
        
        gap = miami_off_pctl - osu_def_qual
        
        if gap >= 15:
            projection = "MIA ▲▲"
        elif gap >= 5:
            projection = "MIA ▲"
        elif gap <= -15:
            projection = "OSU ▲▲"
        elif gap <= -5:
            projection = "OSU ▲"
        else:
            projection = "EVEN"
        
        label = clean_metric_name(off_m).replace('Offense ', '')[:20]
        table_data.append([label, f"{miami_off_pctl:.0f}th", f"{osu_def_qual:.0f}th", f"{gap:+.0f}", projection])
    
    if len(table_data) > 1:
        table = Table(table_data, colWidths=[1.5*inch, 0.9*inch, 0.9*inch, 0.6*inch, 0.9*inch])
        style_cmds = [
            ('BACKGROUND', (0, 0), (-1, 0), MIAMI_GREEN),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]
        for i in range(1, len(table_data)):
            proj = table_data[i][4]
            if "MIA" in proj:
                style_cmds.append(('TEXTCOLOR', (4, i), (4, i), MIAMI_GREEN))
            elif "OSU" in proj:
                style_cmds.append(('TEXTCOLOR', (4, i), (4, i), OSU_SCARLET))
            if i % 2 == 0:
                style_cmds.append(('BACKGROUND', (0, i), (-1, i), LIGHT_GRAY))
        table.setStyle(TableStyle(style_cmds))
        elements.append(table)
    
    elements.append(Spacer(1, 8))
    
    # OSU Offense vs Miami Defense
    elements.append(Paragraph("Ohio State Offense vs Miami Defense", styles['TeamOSU']))
    
    table_data2 = [["MATCHUP", "OSU OFF %", "MIA DEF %", "GAP", "PROJ"]]
    
    for off_m, def_m in zip(offense_metrics, defense_metrics):
        off_key = f"{off_m}_season_avg"
        def_key = f"{def_m}_season_avg"
        
        osu_off_row = comp_df[comp_df['Metric'] == off_key]
        miami_def_row = comp_df[comp_df['Metric'] == def_key]
        
        if osu_off_row.empty or miami_def_row.empty:
            continue
        
        osu_off_pctl = osu_off_row['OSU_Pctl'].values[0]
        miami_def_raw = miami_def_row['Miami_Pctl'].values[0]
        miami_def_qual = convert_to_quality_percentile(miami_def_raw, def_key)
        
        gap = osu_off_pctl - miami_def_qual
        
        if gap >= 15:
            projection = "OSU ▲▲"
        elif gap >= 5:
            projection = "OSU ▲"
        elif gap <= -15:
            projection = "MIA ▲▲"
        elif gap <= -5:
            projection = "MIA ▲"
        else:
            projection = "EVEN"
        
        label = clean_metric_name(off_m).replace('Offense ', '')[:20]
        table_data2.append([label, f"{osu_off_pctl:.0f}th", f"{miami_def_qual:.0f}th", f"{gap:+.0f}", projection])
    
    if len(table_data2) > 1:
        table2 = Table(table_data2, colWidths=[1.5*inch, 0.9*inch, 0.9*inch, 0.6*inch, 0.9*inch])
        style_cmds2 = [
            ('BACKGROUND', (0, 0), (-1, 0), OSU_SCARLET),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]
        for i in range(1, len(table_data2)):
            proj = table_data2[i][4]
            if "MIA" in proj:
                style_cmds2.append(('TEXTCOLOR', (4, i), (4, i), MIAMI_GREEN))
            elif "OSU" in proj:
                style_cmds2.append(('TEXTCOLOR', (4, i), (4, i), OSU_SCARLET))
            if i % 2 == 0:
                style_cmds2.append(('BACKGROUND', (0, i), (-1, i), LIGHT_GRAY))
        table2.setStyle(TableStyle(style_cmds2))
        elements.append(table2)
    
    return elements

def create_matchup_edges(styles, data):
    elements = []
    
    elements.append(Paragraph("KEY MATCHUP EDGES (Strength vs Weakness)", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 5))
    
    edges_df = data['edges']
    
    if edges_df.empty:
        elements.append(Paragraph("No clear strength-vs-weakness matchup edges identified.", styles['Body']))
        return elements
    
    table_data = [["MATCHUP AREA", "EDGE FOR", "STR %", "WEAK %"]]
    
    for _, row in edges_df.iterrows():
        table_data.append([
            clean_metric_name(row['Matchup'])[:28],
            row['Edge_For'],
            f"{row['Team1_Pctl']:.0f}th",
            f"{row['Team2_Pctl']:.0f}th"
        ])
    
    table = Table(table_data, colWidths=[2.4*inch, 1.3*inch, 0.9*inch, 0.9*inch])
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]
    
    for i in range(1, len(table_data)):
        edge_for = table_data[i][1]
        if "Miami" in edge_for:
            style_commands.append(('TEXTCOLOR', (1, i), (1, i), MIAMI_GREEN))
            style_commands.append(('FONTNAME', (1, i), (1, i), 'Helvetica-Bold'))
        elif "Ohio" in edge_for:
            style_commands.append(('TEXTCOLOR', (1, i), (1, i), OSU_SCARLET))
            style_commands.append(('FONTNAME', (1, i), (1, i), 'Helvetica-Bold'))
        if i % 2 == 0:
            style_commands.append(('BACKGROUND', (0, i), (-1, i), LIGHT_GRAY))
    
    table.setStyle(TableStyle(style_commands))
    elements.append(table)
    
    # Add explanatory commentary
    elements.append(Spacer(1, 8))
    
    edge_commentary = """
    
    <b>Red Zone Battle:</b> Ohio State's red zone defense (70th percentile) faces Miami's struggling 
    red zone offense (39th percentile TD rate). When Miami gets inside the 20, they have historically 
    settled for field goals. OSU's ability to force field goal attempts in the red zone could be 
    the difference in a close game - every TD prevented is a 4-point swing.
    """
    
    elements.append(Paragraph(edge_commentary, styles['Body']))
    
    return elements

def create_strengths_vulnerabilities(styles, data):
    elements = []
    
    # VULNERABILITIES
    elements.append(Paragraph("TEAM VULNERABILITIES (Bottom 40%)", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 5))
    
    vuln_df = data['vulnerabilities']
    
    for team, team_color, team_style in [("Miami", MIAMI_GREEN, 'TeamMiami'), 
                                          ("Ohio State", OSU_SCARLET, 'TeamOSU')]:
        elements.append(Paragraph(f"{team} Vulnerabilities", styles[team_style]))
        
        team_vuln = vuln_df[vuln_df['Team'] == team] if not vuln_df.empty else pd.DataFrame()
        
        if team_vuln.empty:
            elements.append(Paragraph("No significant vulnerabilities identified.", styles['Body']))
        else:
            table_data = [["METRIC", "VALUE", "RAW %", "QUAL %"]]
            
            for _, row in team_vuln.iterrows():
                raw_pctl = row['Percentile']
                qual_pctl = convert_to_quality_percentile(raw_pctl, row['Metric'])
                
                table_data.append([
                    clean_metric_name(row['Metric'])[:32],
                    f"{row['Value']:.3f}",
                    f"{raw_pctl:.0f}th",
                    format_percentile(qual_pctl)
                ])
            
            table = Table(table_data, colWidths=[2.6*inch, 0.8*inch, 0.7*inch, 0.8*inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), team_color),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 8),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]))
            elements.append(table)
        
        elements.append(Spacer(1, 6))
    
    # STRENGTHS
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("TEAM STRENGTHS (Top 30%)", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 5))
    
    str_df = data['strengths']
    
    for team, team_color, team_style in [("Miami", MIAMI_GREEN, 'TeamMiami'), 
                                          ("Ohio State", OSU_SCARLET, 'TeamOSU')]:
        elements.append(Paragraph(f"{team} Strengths", styles[team_style]))
        
        team_str = str_df[str_df['Team'] == team] if not str_df.empty else pd.DataFrame()
        
        if team_str.empty:
            elements.append(Paragraph("No significant strengths identified.", styles['Body']))
        else:
            table_data = [["METRIC", "VALUE", "RAW %", "QUAL %"]]
            
            for _, row in team_str.iterrows():
                raw_pctl = row['Percentile']
                qual_pctl = convert_to_quality_percentile(raw_pctl, row['Metric'])
                
                table_data.append([
                    clean_metric_name(row['Metric'])[:32],
                    f"{row['Value']:.3f}",
                    f"{raw_pctl:.0f}th",
                    format_percentile(qual_pctl)
                ])
            
            table = Table(table_data, colWidths=[2.6*inch, 0.8*inch, 0.7*inch, 0.8*inch])
            style_cmds = [
                ('BACKGROUND', (0, 0), (-1, 0), team_color),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 8),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]
            for i in range(1, len(table_data)):
                if i % 2 == 0:
                    style_cmds.append(('BACKGROUND', (0, i), (-1, i), LIGHT_GRAY))
            table.setStyle(TableStyle(style_cmds))
            elements.append(table)
        
        elements.append(Spacer(1, 6))
    
    return elements

def create_trend_analysis(styles, data):
    elements = []
    
    elements.append(Paragraph("MOMENTUM / TREND ANALYSIS (Last 3 vs Season)", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 5))
    
    trends_df = data['trends']
    
    if trends_df.empty:
        elements.append(Paragraph("Trend data not available.", styles['Body']))
        return elements
    
    for team, team_color, team_style in [("Miami", MIAMI_GREEN, 'TeamMiami'), 
                                          ("Ohio State", OSU_SCARLET, 'TeamOSU')]:
        elements.append(Paragraph(f"{team} Trends", styles[team_style]))
        
        team_trends = trends_df[trends_df['Team'] == team]
        
        if team_trends.empty:
            elements.append(Paragraph("No trend data.", styles['Body']))
            continue
        
        table_data = [["METRIC", "SEASON", "LAST 3", "TREND", "DIR"]]
        
        for _, row in team_trends.iterrows():
            if pd.isna(row.get('Trend_3', None)):
                continue
            
            direction = row.get('Direction', 'N/A')
            dir_symbol = "↑" if direction == "IMPROVING" else "↓" if direction == "DECLINING" else "→"
            
            table_data.append([
                clean_metric_name(row['Metric'])[:24],
                f"{row['Season_Avg']:.3f}",
                f"{row['Last_3_Avg']:.3f}",
                f"{row['Trend_3']:+.3f}",
                dir_symbol
            ])
        
        if len(table_data) > 1:
            table = Table(table_data, colWidths=[1.9*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.5*inch])
            style_cmds = [
                ('BACKGROUND', (0, 0), (-1, 0), team_color),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 8),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]
            
            for i in range(1, len(table_data)):
                dir_val = table_data[i][4]
                if dir_val == "↑":
                    style_cmds.append(('TEXTCOLOR', (4, i), (4, i), GREEN_GOOD))
                elif dir_val == "↓":
                    style_cmds.append(('TEXTCOLOR', (4, i), (4, i), RED_BAD))
                if i % 2 == 0:
                    style_cmds.append(('BACKGROUND', (0, i), (-1, i), LIGHT_GRAY))
            
            table.setStyle(TableStyle(style_cmds))
            elements.append(table)
        
        elements.append(Spacer(1, 6))
    
    return elements

def create_quarter_analysis(styles, data):
    elements = []
    
    elements.append(Paragraph("QUARTER-BY-QUARTER PERFORMANCE", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    quarters_df = data['quarters']
    
    if quarters_df.empty:
        elements.append(Paragraph("Quarter analysis data not available.", styles['Body']))
        return elements
    
    table_data = [["QTR", "MIA EPA", "MIA SR%", "MIA PASS%", "OSU EPA", "OSU SR%", "OSU PASS%", "EDGE"]]
    
    for period in sorted(quarters_df['period'].unique()):
        miami_q = quarters_df[(quarters_df['team'] == 'Miami') & (quarters_df['period'] == period)]
        osu_q = quarters_df[(quarters_df['team'] == 'Ohio State') & (quarters_df['period'] == period)]
        
        if miami_q.empty or osu_q.empty:
            continue
        
        miami_epa = miami_q['avg_epa_per_play'].values[0]
        miami_success = miami_q['avg_success_rate'].values[0]
        miami_pass = miami_q['avg_pass_rate'].values[0] if 'avg_pass_rate' in miami_q.columns else 0
        osu_epa = osu_q['avg_epa_per_play'].values[0]
        osu_success = osu_q['avg_success_rate'].values[0]
        osu_pass = osu_q['avg_pass_rate'].values[0] if 'avg_pass_rate' in osu_q.columns else 0
        
        if osu_epa > miami_epa + 0.03:
            edge = "OSU"
        elif miami_epa > osu_epa + 0.03:
            edge = "MIA"
        else:
            edge = "EVEN"
        
        table_data.append([
            f"Q{int(period)}",
            f"{miami_epa:.3f}",
            f"{miami_success*100:.1f}%",
            f"{miami_pass*100:.0f}%",
            f"{osu_epa:.3f}",
            f"{osu_success*100:.1f}%",
            f"{osu_pass*100:.0f}%",
            edge
        ])
    
    # Add half summaries
    miami_1h = quarters_df[(quarters_df['team'] == 'Miami') & (quarters_df['period'] <= 2)]['avg_epa_per_play'].mean()
    miami_2h = quarters_df[(quarters_df['team'] == 'Miami') & (quarters_df['period'] > 2)]['avg_epa_per_play'].mean()
    osu_1h = quarters_df[(quarters_df['team'] == 'Ohio State') & (quarters_df['period'] <= 2)]['avg_epa_per_play'].mean()
    osu_2h = quarters_df[(quarters_df['team'] == 'Ohio State') & (quarters_df['period'] > 2)]['avg_epa_per_play'].mean()
    
    table_data.append(["1H", f"{miami_1h:.3f}", "-", "-", f"{osu_1h:.3f}", "-", "-", 
                       "OSU" if osu_1h > miami_1h else "MIA"])
    table_data.append(["2H", f"{miami_2h:.3f}", "-", "-", f"{osu_2h:.3f}", "-", "-",
                       "OSU" if osu_2h > miami_2h else "MIA"])
    table_data.append(["ADJ", f"{miami_2h - miami_1h:+.3f}", "-", "-", f"{osu_2h - osu_1h:+.3f}", "-", "-", "-"])
    
    table = Table(table_data, colWidths=[0.5*inch, 0.8*inch, 0.7*inch, 0.75*inch, 0.8*inch, 0.7*inch, 0.75*inch, 0.6*inch])
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 7),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('BACKGROUND', (0, -3), (-1, -1), colors.HexColor("#e8e8e8")),
    ]
    
    for i in range(1, len(table_data)):
        edge = table_data[i][7]
        if "MIA" in str(edge):
            style_commands.append(('TEXTCOLOR', (7, i), (7, i), MIAMI_GREEN))
        elif "OSU" in str(edge):
            style_commands.append(('TEXTCOLOR', (7, i), (7, i), OSU_SCARLET))
    
    table.setStyle(TableStyle(style_commands))
    elements.append(table)
    
    return elements

def create_situational_analysis(styles, data):
    elements = []
    
    elements.append(Paragraph("DOWN & DISTANCE SITUATIONAL ANALYSIS", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    sit_df = data['situations']
    
    if sit_df.empty:
        elements.append(Paragraph("Situational data not available.", styles['Body']))
        return elements
    
    situations_order = ["1st Down", "2nd & Short (1-3)", "2nd & Medium (4-6)", "2nd & Long (7+)",
                        "3rd & Short (1-3)", "3rd & Medium (4-6)", "3rd & Long (7+)", "4th Down"]
    
    table_data = [["SITUATION", "MIA EPA", "MIA SR%", "OSU EPA", "OSU SR%", "EDGE"]]
    
    for sit in situations_order:
        miami_sit = sit_df[(sit_df['team'] == 'Miami') & (sit_df['situation'] == sit)]
        osu_sit = sit_df[(sit_df['team'] == 'Ohio State') & (sit_df['situation'] == sit)]
        
        if miami_sit.empty or osu_sit.empty:
            continue
        
        miami_epa = miami_sit['epa_per_play'].values[0]
        miami_sr = miami_sit['success_rate'].values[0]
        osu_epa = osu_sit['epa_per_play'].values[0]
        osu_sr = osu_sit['success_rate'].values[0]
        
        if osu_epa > miami_epa + 0.05:
            edge = "OSU"
        elif miami_epa > osu_epa + 0.05:
            edge = "MIA"
        else:
            edge = "EVEN"
        
        table_data.append([
            sit[:20],
            f"{miami_epa:.2f}",
            f"{miami_sr*100:.0f}%",
            f"{osu_epa:.2f}",
            f"{osu_sr*100:.0f}%",
            edge
        ])
    
    if len(table_data) > 1:
        table = Table(table_data, colWidths=[1.6*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.7*inch])
        
        style_commands = [
            ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]
        
        for i in range(1, len(table_data)):
            edge = table_data[i][5]
            if "MIA" in edge:
                style_commands.append(('TEXTCOLOR', (5, i), (5, i), MIAMI_GREEN))
            elif "OSU" in edge:
                style_commands.append(('TEXTCOLOR', (5, i), (5, i), OSU_SCARLET))
            if i % 2 == 0:
                style_commands.append(('BACKGROUND', (0, i), (-1, i), LIGHT_GRAY))
        
        table.setStyle(TableStyle(style_commands))
        elements.append(table)
    
    return elements

def create_field_position_analysis(styles, data):
    elements = []
    
    elements.append(Paragraph("FIELD POSITION PERFORMANCE", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    fp_df = data['field_position']
    
    if fp_df.empty:
        elements.append(Paragraph("Field position data not available.", styles['Body']))
        return elements
    
    zones_order = ["Own 0-20 (Backed Up)", "Own 21-40", "Midfield (41-60)", 
                   "Opponent 21-40", "Red Zone (Opp 0-20)"]
    
    table_data = [["FIELD ZONE", "MIA EPA", "MIA SR%", "OSU EPA", "OSU SR%", "EDGE"]]
    
    for zone in zones_order:
        miami_fp = fp_df[(fp_df['team'] == 'Miami') & (fp_df['field_zone'] == zone)]
        osu_fp = fp_df[(fp_df['team'] == 'Ohio State') & (fp_df['field_zone'] == zone)]
        
        if miami_fp.empty or osu_fp.empty:
            continue
        
        miami_epa = miami_fp['epa_per_play'].values[0]
        miami_sr = miami_fp['success_rate'].values[0]
        osu_epa = osu_fp['epa_per_play'].values[0]
        osu_sr = osu_fp['success_rate'].values[0]
        
        if osu_epa > miami_epa + 0.05:
            edge = "OSU"
        elif miami_epa > osu_epa + 0.05:
            edge = "MIA"
        else:
            edge = "EVEN"
        
        table_data.append([
            zone[:22],
            f"{miami_epa:.2f}",
            f"{miami_sr*100:.0f}%",
            f"{osu_epa:.2f}",
            f"{osu_sr*100:.0f}%",
            edge
        ])
    
    if len(table_data) > 1:
        table = Table(table_data, colWidths=[1.7*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.6*inch])
        
        style_commands = [
            ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]
        
        for i in range(1, len(table_data)):
            edge = table_data[i][5]
            if "MIA" in edge:
                style_commands.append(('TEXTCOLOR', (5, i), (5, i), MIAMI_GREEN))
            elif "OSU" in edge:
                style_commands.append(('TEXTCOLOR', (5, i), (5, i), OSU_SCARLET))
            if i % 2 == 0:
                style_commands.append(('BACKGROUND', (0, i), (-1, i), LIGHT_GRAY))
        
        table.setStyle(TableStyle(style_commands))
        elements.append(table)
    
    return elements

def create_drive_efficiency(styles, data):
    elements = []
    
    elements.append(Paragraph("DRIVE EFFICIENCY COMPARISON", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    drives_df = data['drives']
    
    if drives_df.empty:
        elements.append(Paragraph("Drive efficiency data not available.", styles['Body']))
        return elements
    
    miami_dr = drives_df[drives_df['team'] == 'Miami']
    osu_dr = drives_df[drives_df['team'] == 'Ohio State']
    
    if miami_dr.empty or osu_dr.empty:
        elements.append(Paragraph("Drive data incomplete.", styles['Body']))
        return elements
    
    metrics = [
        ('avg_plays_per_drive', 'Plays/Drive'),
        ('avg_yards_per_drive', 'Yards/Drive'),
        ('avg_epa_per_drive', 'EPA/Drive'),
        ('scoring_pct', 'Scoring %'),
        ('points_per_drive', 'Pts/Drive'),
        ('td_pct', 'TD %'),
        ('avg_start_position', 'Avg Start Yard')
    ]
    
    table_data = [["METRIC", "MIAMI", "OHIO STATE", "EDGE"]]
    
    for col, label in metrics:
        if col not in miami_dr.columns or col not in osu_dr.columns:
            continue
        
        miami_val = miami_dr[col].values[0]
        osu_val = osu_dr[col].values[0]
        
        # Determine edge
        if col in ['scoring_pct', 'td_pct']:
            edge = "MIA" if miami_val > osu_val + 0.05 else "OSU" if osu_val > miami_val + 0.05 else "EVEN"
        elif col == 'avg_start_position':
            edge = "MIA" if miami_val > osu_val + 5 else "OSU" if osu_val > miami_val + 5 else "EVEN"
        else:
            edge = "MIA" if miami_val > osu_val * 1.05 else "OSU" if osu_val > miami_val * 1.05 else "EVEN"
        
        if col in ['scoring_pct', 'td_pct']:
            table_data.append([label, f"{miami_val*100:.1f}%", f"{osu_val*100:.1f}%", edge])
        else:
            table_data.append([label, f"{miami_val:.2f}", f"{osu_val:.2f}", edge])
    
    table = Table(table_data, colWidths=[1.8*inch, 1.2*inch, 1.2*inch, 0.8*inch])
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]
    
    for i in range(1, len(table_data)):
        edge = table_data[i][3]
        if "MIA" in edge:
            style_commands.append(('TEXTCOLOR', (3, i), (3, i), MIAMI_GREEN))
        elif "OSU" in edge:
            style_commands.append(('TEXTCOLOR', (3, i), (3, i), OSU_SCARLET))
        if i % 2 == 0:
            style_commands.append(('BACKGROUND', (0, i), (-1, i), LIGHT_GRAY))
    
    table.setStyle(TableStyle(style_commands))
    elements.append(table)
    
    return elements

def create_pressure_analysis(styles, data):
    elements = []
    
    elements.append(Paragraph("PRESSURE & PROTECTION ANALYSIS", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    pressure_df = data['pressure']
    
    if pressure_df.empty:
        elements.append(Paragraph("Pressure analysis data not available.", styles['Body']))
        return elements
    
    # Protection (offense) - uses different columns
    elements.append(Paragraph("Pass Protection (Offense)", styles['SubSection']))
    
    miami_prot = pressure_df[(pressure_df['team'] == 'Miami') & (pressure_df['side'] == 'Protection')]
    osu_prot = pressure_df[(pressure_df['team'] == 'Ohio State') & (pressure_df['side'] == 'Protection')]
    
    if not miami_prot.empty and not osu_prot.empty:
        table_data = [["METRIC", "MIAMI", "OHIO STATE"]]
        
        metrics = [
            ('total_dropbacks', 'Total Dropbacks', False),  # count
            ('sacks_allowed', 'Sacks Allowed', False),  # count
            ('sack_rate_allowed', 'Sack Rate Allowed', True),  # rate
            ('negative_pass_rate', 'Negative Pass Rate', True)  # rate
        ]
        
        for col, label, is_rate in metrics:
            if col in miami_prot.columns:
                m_val = miami_prot[col].values[0]
                o_val = osu_prot[col].values[0]
                
                if pd.notna(m_val) and pd.notna(o_val):
                    if is_rate:
                        table_data.append([label, f"{m_val*100:.1f}%", f"{o_val*100:.1f}%"])
                    else:
                        table_data.append([label, f"{int(m_val)}", f"{int(o_val)}"])
        
        if len(table_data) > 1:
            table = Table(table_data, colWidths=[2*inch, 1.2*inch, 1.2*inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 8),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ]))
            elements.append(table)
        else:
            elements.append(Paragraph("Protection data not available.", styles['Body']))
    else:
        elements.append(Paragraph("Protection data not available.", styles['Body']))
    
    elements.append(Spacer(1, 8))
    
    # Pass Rush (defense) - uses different columns
    elements.append(Paragraph("Pass Rush (Defense)", styles['SubSection']))
    
    miami_rush = pressure_df[(pressure_df['team'] == 'Miami') & (pressure_df['side'] == 'Pass Rush')]
    osu_rush = pressure_df[(pressure_df['team'] == 'Ohio State') & (pressure_df['side'] == 'Pass Rush')]
    
    if not miami_rush.empty and not osu_rush.empty:
        table_data = [["METRIC", "MIAMI", "OHIO STATE"]]
        
        metrics = [
            ('total_dropbacks_faced', 'Dropbacks Faced', False),  # count
            ('sacks_generated', 'Sacks Generated', False),  # count
            ('sack_rate_generated', 'Sack Rate', True),  # rate
            ('disruption_rate', 'Disruption Rate', True)  # rate
        ]
        
        for col, label, is_rate in metrics:
            if col in miami_rush.columns:
                m_val = miami_rush[col].values[0]
                o_val = osu_rush[col].values[0]
                
                if pd.notna(m_val) and pd.notna(o_val):
                    if is_rate:
                        table_data.append([label, f"{m_val*100:.1f}%", f"{o_val*100:.1f}%"])
                    else:
                        table_data.append([label, f"{int(m_val)}", f"{int(o_val)}"])
        
        if len(table_data) > 1:
            table = Table(table_data, colWidths=[2*inch, 1.2*inch, 1.2*inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 8),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ]))
            elements.append(table)
        else:
            elements.append(Paragraph("Pass rush data not available.", styles['Body']))
    else:
        elements.append(Paragraph("Pass rush data not available.", styles['Body']))
    
    return elements

def create_game_script(styles, data):
    elements = []
    
    elements.append(Paragraph("GAME SCRIPT PERFORMANCE (By Score Differential)", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    script_df = data['game_script']
    
    if script_df.empty:
        elements.append(Paragraph("Game script data not available.", styles['Body']))
        return elements
    
    for team, team_color, team_style in [("Miami", MIAMI_GREEN, 'TeamMiami'), 
                                          ("Ohio State", OSU_SCARLET, 'TeamOSU')]:
        elements.append(Paragraph(f"{team}", styles[team_style]))
        
        team_script = script_df[script_df['Team'] == team]
        
        if team_script.empty:
            elements.append(Paragraph("No game script data.", styles['Body']))
            continue
        
        table_data = [["SITUATION", "PLAYS", "EPA/PLAY", "SR%", "PASS%"]]
        
        for _, row in team_script.iterrows():
            table_data.append([
                row['game_script'].replace('_', ' '),
                f"{int(row['Plays'])}",
                f"{row['EPA_per_Play']:.3f}",
                f"{row['Success_Rate']*100:.1f}%",
                f"{row['Pass_Rate']*100:.1f}%"
            ])
        
        table = Table(table_data, colWidths=[1.5*inch, 0.7*inch, 0.9*inch, 0.7*inch, 0.7*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), team_color),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 6))
    
    return elements

def create_late_close_analysis(styles, data):
    elements = []
    
    elements.append(Paragraph("LATE & CLOSE GAME PERFORMANCE (Q4 + Within 8 pts)", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    lc_df = data['late_close']
    
    if lc_df.empty:
        elements.append(Paragraph("Late/close game data not available.", styles['Body']))
        return elements
    
    table_data = [["TEAM", "SITUATION", "PLAYS", "EPA/PLAY", "SR%", "EXPL%"]]
    
    for team in ["Miami", "Ohio State"]:
        team_lc = lc_df[lc_df['team'] == team]
        for _, row in team_lc.iterrows():
            table_data.append([
                team[:6],
                row['situation'][:18],
                f"{int(row['plays'])}",
                f"{row['epa_per_play']:.3f}",
                f"{row['success_rate']*100:.1f}%",
                f"{row['explosive_rate']*100:.1f}%"
            ])
    
    table = Table(table_data, colWidths=[0.8*inch, 1.6*inch, 0.7*inch, 0.9*inch, 0.7*inch, 0.7*inch])
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 1), (1, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]
    
    for i in range(1, len(table_data)):
        team = table_data[i][0]
        if "Miami" in team:
            style_commands.append(('TEXTCOLOR', (0, i), (0, i), MIAMI_GREEN))
        elif "Ohio" in team:
            style_commands.append(('TEXTCOLOR', (0, i), (0, i), OSU_SCARLET))
    
    table.setStyle(TableStyle(style_commands))
    elements.append(table)
    
    return elements

def create_tempo_analysis(styles, data):
    elements = []
    
    elements.append(Paragraph("TEMPO & PACE ANALYSIS", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    tempo_df = data['tempo']
    
    if tempo_df.empty:
        elements.append(Paragraph("Tempo data not available.", styles['Body']))
        return elements
    
    table_data = [["TEAM", "AVG PLAYS/GAME", "MIN", "MAX", "STD DEV"]]
    
    for _, row in tempo_df.iterrows():
        table_data.append([
            row['team'],
            f"{row['avg_plays_per_game']:.1f}",
            f"{int(row['min_plays'])}",
            f"{int(row['max_plays'])}",
            f"{row['sd_plays']:.1f}"
        ])
    
    table = Table(table_data, colWidths=[1.5*inch, 1.2*inch, 0.8*inch, 0.8*inch, 0.9*inch])
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]
    
    table.setStyle(TableStyle(style_commands))
    elements.append(table)
    
    return elements

# ============================================================================
# LINE OF SCRIMMAGE ANALYSIS SECTIONS
# ============================================================================

def create_los_summary(styles, data):
    """Overall Line of Scrimmage Summary"""
    elements = []
    
    elements.append(Paragraph("LINE OF SCRIMMAGE SUMMARY", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    los_df = data['los_summary']
    
    if los_df.empty:
        elements.append(Paragraph("LOS summary data not available.", styles['Body']))
        return elements
    
    # O-Line comparison
    elements.append(Paragraph("Offensive Line Performance", styles['SubSection']))
    
    miami_oline = los_df[(los_df['team'] == 'Miami') & (los_df['side'] == 'O-Line')]
    osu_oline = los_df[(los_df['team'] == 'Ohio State') & (los_df['side'] == 'O-Line')]
    
    if not miami_oline.empty and not osu_oline.empty:
        table_data = [["METRIC", "MIAMI", "OHIO STATE", "EDGE"]]
        
        metrics = [
            ('rush_ypc', 'Run Block: YPC'),
            ('rush_success_rate', 'Run Block: Success %'),
            ('stuff_rate_faced', 'Stuffed Rate'),
            ('explosive_rush_rate', 'Explosive Rush %'),
            ('sack_rate', 'Sacks Allowed %'),
            ('negative_pass_rate', 'Negative Pass %')
        ]
        
        for col, label in metrics:
            if col in miami_oline.columns:
                m_val = miami_oline[col].values[0]
                o_val = osu_oline[col].values[0]
                
                # Determine edge (lower is better for stuff_rate, sack_rate, negative rates)
                if 'stuff' in col or 'sack' in col or 'negative' in col:
                    edge = "MIA" if m_val < o_val - 0.01 else "OSU" if o_val < m_val - 0.01 else "EVEN"
                else:
                    edge = "MIA" if m_val > o_val + 0.01 else "OSU" if o_val > m_val + 0.01 else "EVEN"
                
                if 'rate' in col or 'pct' in col.lower():
                    table_data.append([label, f"{m_val*100:.1f}%", f"{o_val*100:.1f}%", edge])
                else:
                    table_data.append([label, f"{m_val:.2f}", f"{o_val:.2f}", edge])
        
        table = Table(table_data, colWidths=[1.8*inch, 1.1*inch, 1.1*inch, 0.7*inch])
        style_cmds = [
            ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]
        for i in range(1, len(table_data)):
            edge = table_data[i][3]
            if "MIA" in edge:
                style_cmds.append(('TEXTCOLOR', (3, i), (3, i), MIAMI_GREEN))
            elif "OSU" in edge:
                style_cmds.append(('TEXTCOLOR', (3, i), (3, i), OSU_SCARLET))
        table.setStyle(TableStyle(style_cmds))
        elements.append(table)
    
    elements.append(Spacer(1, 8))
    
    # D-Line comparison
    elements.append(Paragraph("Defensive Line Performance", styles['SubSection']))
    
    miami_dline = los_df[(los_df['team'] == 'Miami') & (los_df['side'] == 'D-Line')]
    osu_dline = los_df[(los_df['team'] == 'Ohio State') & (los_df['side'] == 'D-Line')]
    
    if not miami_dline.empty and not osu_dline.empty:
        table_data = [["METRIC", "MIAMI", "OHIO STATE", "EDGE"]]
        
        metrics = [
            ('rush_ypc_allowed', 'YPC Allowed'),
            ('rush_success_allowed', 'Rush Success Allowed'),
            ('stuff_rate', 'Stuff Rate'),
            ('tfl_rate', 'TFL Rate'),
            ('sack_rate', 'Sack Rate'),
            ('disruption_rate', 'Disruption Rate')
        ]
        
        for col, label in metrics:
            if col in miami_dline.columns:
                m_val = miami_dline[col].values[0]
                o_val = osu_dline[col].values[0]
                
                # Lower is better for allowed metrics, higher for stuff/sack/disruption/tfl
                if 'allowed' in col or 'ypc' in col:
                    edge = "MIA" if m_val < o_val - 0.01 else "OSU" if o_val < m_val - 0.01 else "EVEN"
                else:
                    edge = "MIA" if m_val > o_val + 0.01 else "OSU" if o_val > m_val + 0.01 else "EVEN"
                
                # YPC is numeric, rates are percentages
                if 'ypc' in col:
                    table_data.append([label, f"{m_val:.2f}", f"{o_val:.2f}", edge])
                elif 'rate' in col or 'success' in col:
                    table_data.append([label, f"{m_val*100:.1f}%", f"{o_val*100:.1f}%", edge])
                else:
                    table_data.append([label, f"{m_val:.2f}", f"{o_val:.2f}", edge])
        
        table = Table(table_data, colWidths=[1.8*inch, 1.1*inch, 1.1*inch, 0.7*inch])
        style_cmds = [
            ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]
        for i in range(1, len(table_data)):
            edge = table_data[i][3]
            if "MIA" in edge:
                style_cmds.append(('TEXTCOLOR', (3, i), (3, i), MIAMI_GREEN))
            elif "OSU" in edge:
                style_cmds.append(('TEXTCOLOR', (3, i), (3, i), OSU_SCARLET))
        table.setStyle(TableStyle(style_cmds))
        elements.append(table)
    
    return elements

def create_run_progression_analysis(styles, data):
    """Run game progression through quarters"""
    elements = []
    
    elements.append(Paragraph("RUN GAME PROGRESSION (O-Line Control)", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    quarters_df = data['run_progression_quarters']
    trend_df = data['run_progression_trend']
    
    if quarters_df.empty:
        elements.append(Paragraph("Run progression data not available.", styles['Body']))
        return elements
    
    # Quarter by quarter for each team
    for team, team_color, team_style in [("Miami", MIAMI_GREEN, 'TeamMiami'), 
                                          ("Ohio State", OSU_SCARLET, 'TeamOSU')]:
        elements.append(Paragraph(f"{team} Run Game by Quarter", styles[team_style]))
        
        team_q = quarters_df[quarters_df['team'] == team]
        
        if team_q.empty:
            continue
        
        table_data = [["QTR", "YPC", "SR%", "STUFF%", "EXPL%", "EPA"]]
        
        for _, row in team_q.sort_values('period').iterrows():
            if row['period'] > 4:
                continue
            table_data.append([
                f"Q{int(row['period'])}",
                f"{row['avg_ypc']:.2f}",
                f"{row['avg_rush_success']*100:.1f}%",
                f"{row['avg_stuff_rate']*100:.1f}%",
                f"{row['avg_explosive_rate']*100:.1f}%",
                f"{row['avg_rush_epa']:.3f}"
            ])
        
        table = Table(table_data, colWidths=[0.5*inch, 0.7*inch, 0.7*inch, 0.8*inch, 0.7*inch, 0.7*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), team_color),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 4))
    
    # Trend Summary
    if not trend_df.empty:
        elements.append(Spacer(1, 6))
        elements.append(Paragraph("Half-to-Half Run Game Trend", styles['SubSection']))
        
        table_data = [["TEAM", "1H SR%", "2H SR%", "TREND", "1H YPC", "2H YPC", "TREND"]]
        
        for _, row in trend_df.iterrows():
            sr_trend = row['rush_success_trend']
            ypc_trend = row['ypc_trend']
            
            sr_arrow = "↑" if sr_trend > 0.02 else "↓" if sr_trend < -0.02 else "→"
            ypc_arrow = "↑" if ypc_trend > 0.2 else "↓" if ypc_trend < -0.2 else "→"
            
            table_data.append([
                row['team'],
                f"{row['first_half_rush_success']*100:.1f}%",
                f"{row['second_half_rush_success']*100:.1f}%",
                f"{sr_trend:+.1%} {sr_arrow}",
                f"{row['first_half_ypc']:.2f}",
                f"{row['second_half_ypc']:.2f}",
                f"{ypc_trend:+.2f} {ypc_arrow}"
            ])
        
        table = Table(table_data, colWidths=[1.1*inch, 0.7*inch, 0.7*inch, 0.9*inch, 0.7*inch, 0.7*inch, 0.9*inch])
        style_cmds = [
            ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 7),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]
        table.setStyle(TableStyle(style_cmds))
        elements.append(table)
    
    return elements

def create_obvious_run_analysis(styles, data):
    """Run game in obvious run situations"""
    elements = []
    
    elements.append(Paragraph("OBVIOUS RUN SITUATIONS (Can They Close?)", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    obvious_df = data['obvious_run']
    
    if obvious_df.empty:
        elements.append(Paragraph("Obvious run data not available.", styles['Body']))
        return elements
    
    for team, team_color, team_style in [("Miami", MIAMI_GREEN, 'TeamMiami'), 
                                          ("Ohio State", OSU_SCARLET, 'TeamOSU')]:
        elements.append(Paragraph(f"{team}", styles[team_style]))
        
        team_data = obvious_df[obvious_df['team'] == team]
        
        if team_data.empty:
            continue
        
        table_data = [["SITUATION", "PLAYS", "RUSH%", "SR%", "YPC", "CONV%"]]
        
        for _, row in team_data.iterrows():
            table_data.append([
                row['situation'][:22],
                f"{int(row['total_plays'])}",
                f"{row['rush_rate']*100:.0f}%",
                f"{row['rush_success']*100:.1f}%" if pd.notna(row['rush_success']) else "-",
                f"{row['avg_yards']:.1f}" if pd.notna(row['avg_yards']) else "-",
                f"{row['conversion_rate']*100:.0f}%" if pd.notna(row['conversion_rate']) else "-"
            ])
        
        table = Table(table_data, colWidths=[1.7*inch, 0.6*inch, 0.7*inch, 0.6*inch, 0.6*inch, 0.7*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), team_color),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 6))
    
    return elements

def create_run_defense_progression(styles, data):
    """Run defense progression through quarters"""
    elements = []
    
    elements.append(Paragraph("RUN DEFENSE PROGRESSION (D-Line Durability)", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    quarters_df = data['run_defense_quarters']
    trend_df = data['run_defense_trend']
    
    if quarters_df.empty:
        elements.append(Paragraph("Run defense progression data not available.", styles['Body']))
        return elements
    
    for team, team_color, team_style in [("Miami", MIAMI_GREEN, 'TeamMiami'), 
                                          ("Ohio State", OSU_SCARLET, 'TeamOSU')]:
        elements.append(Paragraph(f"{team} Run Defense by Quarter", styles[team_style]))
        
        team_q = quarters_df[quarters_df['team'] == team]
        
        if team_q.empty:
            continue
        
        table_data = [["QTR", "YPC", "SR%", "STUFF%", "TFL%", "EPA"]]
        
        for _, row in team_q.sort_values('period').iterrows():
            if row['period'] > 4:
                continue
            table_data.append([
                f"Q{int(row['period'])}",
                f"{row['avg_ypc_allowed']:.2f}",
                f"{row['avg_rush_success_allowed']*100:.1f}%",
                f"{row['avg_stuff_rate']*100:.1f}%",
                f"{row['avg_tfl_rate']*100:.1f}%",
                f"{row['avg_rush_epa_allowed']:.3f}"
            ])
        
        table = Table(table_data, colWidths=[0.5*inch, 0.7*inch, 0.7*inch, 0.8*inch, 0.7*inch, 0.7*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), team_color),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 4))
    
    # Trend Summary
    if not trend_df.empty:
        elements.append(Spacer(1, 6))
        elements.append(Paragraph("Half-to-Half Run Defense Trend", styles['SubSection']))
        
        table_data = [["TEAM", "1H STUFF%", "2H STUFF%", "TREND", "1H YPC", "2H YPC", "TREND"]]
        
        for _, row in trend_df.iterrows():
            stuff_trend = row['stuff_rate_trend']
            ypc_trend = row['ypc_allowed_trend']
            
            # For defense: improving = higher stuff rate, lower YPC allowed
            stuff_arrow = "↑" if stuff_trend > 0.02 else "↓" if stuff_trend < -0.02 else "→"
            ypc_arrow = "↑" if ypc_trend > 0.2 else "↓" if ypc_trend < -0.2 else "→"
            
            table_data.append([
                row['team'],
                f"{row['first_half_stuff_rate']*100:.1f}%",
                f"{row['second_half_stuff_rate']*100:.1f}%",
                f"{stuff_trend:+.1%} {stuff_arrow}",
                f"{row['first_half_ypc_allowed']:.2f}",
                f"{row['second_half_ypc_allowed']:.2f}",
                f"{ypc_trend:+.2f} {ypc_arrow}"
            ])
        
        table = Table(table_data, colWidths=[1.1*inch, 0.8*inch, 0.8*inch, 0.9*inch, 0.7*inch, 0.7*inch, 0.9*inch])
        style_cmds = [
            ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 7),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]
        table.setStyle(TableStyle(style_cmds))
        elements.append(table)
    
    return elements

def create_run_defense_obvious(styles, data):
    """Run defense in obvious run situations"""
    elements = []
    
    elements.append(Paragraph("RUN DEFENSE IN OBVIOUS RUN SITUATIONS", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    obvious_df = data['run_defense_obvious']
    
    if obvious_df.empty:
        elements.append(Paragraph("Run defense obvious data not available.", styles['Body']))
        return elements
    
    for team, team_color, team_style in [("Miami", MIAMI_GREEN, 'TeamMiami'), 
                                          ("Ohio State", OSU_SCARLET, 'TeamOSU')]:
        elements.append(Paragraph(f"{team}", styles[team_style]))
        
        team_data = obvious_df[obvious_df['team'] == team]
        
        if team_data.empty:
            continue
        
        table_data = [["SITUATION", "PLAYS", "STUFF%", "SR ALWD", "YPC", "EPA"]]
        
        for _, row in team_data.iterrows():
            table_data.append([
                row['situation'][:22],
                f"{int(row['plays'])}",
                f"{row['stuff_rate']*100:.1f}%",
                f"{row['success_allowed']*100:.1f}%",
                f"{row['avg_yards_allowed']:.1f}",
                f"{row['epa_allowed']:.3f}"
            ])
        
        table = Table(table_data, colWidths=[1.7*inch, 0.6*inch, 0.7*inch, 0.8*inch, 0.6*inch, 0.7*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), team_color),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 6))
    
    return elements

def create_protection_obvious_analysis(styles, data):
    """Pass protection in obvious pass situations"""
    elements = []
    
    elements.append(Paragraph("PASS PROTECTION IN OBVIOUS PASS SITUATIONS", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    prot_df = data['protection_obvious']
    
    if prot_df.empty:
        elements.append(Paragraph("Protection obvious pass data not available.", styles['Body']))
        return elements
    
    for team, team_color, team_style in [("Miami", MIAMI_GREEN, 'TeamMiami'), 
                                          ("Ohio State", OSU_SCARLET, 'TeamOSU')]:
        elements.append(Paragraph(f"{team}", styles[team_style]))
        
        team_data = prot_df[prot_df['team'] == team]
        
        if team_data.empty:
            continue
        
        table_data = [["SITUATION", "PLAYS", "SACK%", "SR%", "CONV%", "EPA"]]
        
        for _, row in team_data.iterrows():
            table_data.append([
                row['situation'][:22],
                f"{int(row['plays'])}",
                f"{row['sack_rate']*100:.1f}%",
                f"{row['success_rate']*100:.1f}%",
                f"{row['conversion_rate']*100:.0f}%",
                f"{row['epa_per_play']:.3f}"
            ])
        
        table = Table(table_data, colWidths=[1.7*inch, 0.6*inch, 0.7*inch, 0.6*inch, 0.7*inch, 0.7*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), team_color),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 6))
    
    return elements

def create_pass_rush_obvious_analysis(styles, data):
    """Pass rush in obvious pass situations"""
    elements = []
    
    elements.append(Paragraph("PASS RUSH IN OBVIOUS PASS SITUATIONS", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    rush_df = data['pass_rush_obvious']
    
    if rush_df.empty:
        elements.append(Paragraph("Pass rush obvious data not available.", styles['Body']))
        return elements
    
    for team, team_color, team_style in [("Miami", MIAMI_GREEN, 'TeamMiami'), 
                                          ("Ohio State", OSU_SCARLET, 'TeamOSU')]:
        elements.append(Paragraph(f"{team}", styles[team_style]))
        
        team_data = rush_df[rush_df['team'] == team]
        
        if team_data.empty:
            continue
        
        table_data = [["SITUATION", "PLAYS", "SACK%", "NEG%", "SR ALWD", "EPA"]]
        
        for _, row in team_data.iterrows():
            table_data.append([
                row['situation'][:22],
                f"{int(row['plays'])}",
                f"{row['sack_rate']*100:.1f}%",
                f"{row['negative_play_rate']*100:.1f}%",
                f"{row['success_allowed']*100:.1f}%",
                f"{row['epa_allowed']:.3f}"
            ])
        
        table = Table(table_data, colWidths=[1.7*inch, 0.6*inch, 0.7*inch, 0.6*inch, 0.8*inch, 0.7*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), team_color),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 6))
    
    return elements

def create_summary_counts(styles, data):
    """Summary of edges and advantages"""
    elements = []
    
    elements.append(Paragraph("MATCHUP SUMMARY", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 6))
    
    comp_df = data['comparison']
    edges_df = data['edges']
    
    # Count advantages
    if not comp_df.empty:
        miami_adv = len(comp_df[comp_df['Advantage'] == 'MIAMI'])
        osu_adv = len(comp_df[comp_df['Advantage'] == 'OHIO STATE'])
        even = len(comp_df[comp_df['Advantage'] == 'EVEN'])
    else:
        miami_adv = osu_adv = even = 0
    
    if not edges_df.empty:
        miami_edges = len(edges_df[edges_df['Edge_For'] == 'Miami'])
        osu_edges = len(edges_df[edges_df['Edge_For'] == 'Ohio State'])
    else:
        miami_edges = osu_edges = 0
    
    table_data = [
        ["CATEGORY", "MIAMI", "OHIO STATE", "EVEN"],
        ["Statistical Advantages", str(miami_adv), str(osu_adv), str(even)],
        ["Strength vs Weakness Edges", str(miami_edges), str(osu_edges), "-"]
    ]
    
    table = Table(table_data, colWidths=[2*inch, 1.2*inch, 1.2*inch, 0.8*inch])
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TEXTCOLOR', (1, 1), (1, -1), MIAMI_GREEN),
        ('TEXTCOLOR', (2, 1), (2, -1), OSU_SCARLET),
        ('FONTNAME', (1, 1), (2, -1), 'Helvetica-Bold'),
    ]
    
    table.setStyle(TableStyle(style_commands))
    elements.append(table)
    
    return elements

# ============================================================================
# COMMENTARY AND ANALYSIS FUNCTIONS
# ============================================================================

def create_executive_summary(styles, data):
    """Executive summary with key findings"""
    elements = []
    
    elements.append(Paragraph("EXECUTIVE SUMMARY", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 8))
    
    comp_df = data['comparison']
    
    # Count advantages
    miami_adv = len(comp_df[comp_df['Advantage'] == 'MIAMI']) if not comp_df.empty else 0
    osu_adv = len(comp_df[comp_df['Advantage'] == 'OHIO STATE']) if not comp_df.empty else 0
    
    summary_text = f"""
    <b>Overall Statistical Edge: Ohio State</b><br/><br/>
    
    Ohio State holds advantages in {osu_adv} of the tracked metrics compared to Miami's {miami_adv}. 
    The Buckeyes enter this matchup as the more complete team statistically, ranking elite (90th+ percentile) 
    in both offensive efficiency and defensive efficiency. Miami's defense has been strong, but the Hurricanes 
    face significant challenges on both sides of the ball against this opponent.<br/><br/>
    
    <b>Key Statistical Storylines:</b><br/>
    • <b>Offensive Disparity:</b> Ohio State's offense ranks 91st percentile in EPA/play vs Miami's 69th - a significant gap<br/>
    • <b>Defensive Parity:</b> Both defenses are elite - OSU 94th percentile, Miami 90th percentile in defensive EPA<br/>
    • <b>Red Zone Concern:</b> Miami's red zone defense is a glaring weakness (7th percentile in TD rate allowed vs OSU's 70th)<br/>
    • <b>Line of Scrimmage:</b> Ohio State's run game improves significantly in 2nd half (+8.6% success rate); Miami's D-line allows +0.76 YPC more<br/>
    • <b>Situational Pass Rush:</b> OSU's 16.4% sack rate on 3rd and long dwarfs Miami's 3.8%<br/>
    """
    
    elements.append(Paragraph(summary_text, styles['Body']))
    
    return elements

def create_offensive_analysis(styles, data):
    """Detailed offensive analysis commentary"""
    elements = []
    
    elements.append(Paragraph("OFFENSIVE ANALYSIS", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 8))
    
    commentary = """
    <b>Ohio State Offense (Elite - 91st Percentile EPA/Play)</b><br/><br/>
    
    The Buckeyes possess one of the nation's most efficient offenses. Their passing attack is devastating - 
    ranking 96th percentile in EPA/pass with a 99th percentile pass success rate. Julian Sayin has been 
    remarkably efficient, completing passes at a 77.3% clip (100th percentile). The offense excels in 
    every down-and-distance situation, particularly on 3rd downs (99th percentile, 55.6% conversion).<br/><br/>
    
    <b>Concern:</b> Recent trends show offensive regression - EPA/play dropped from 0.202 season average 
    to 0.096 in last 3 games. Rush EPA has gone negative (-0.061). The offense may be hitting a late-season lull.
    The offensive line gets stuffed at a higher rate (18.8%) than Miami's (14.3%).<br/><br/>
    
    <b>Miami Offense (Good - 69th Percentile EPA/Play)</b><br/><br/>
    
    Miami's offense is solid but not spectacular. Carson Beck leads an efficient passing game (82nd percentile 
    EPA/pass) with elite completion percentage (73.1%, 99th percentile). The Hurricanes excel in pass success 
    rate (90th percentile). The offensive line actually does well avoiding negative plays - only 14.3% stuffed 
    rate compared to OSU's 18.8%.<br/><br/>
    
    <b>Major Concerns:</b><br/>
    • Run game is below average (42nd percentile EPA/rush) - can't establish ground game despite good pass protection<br/>
    • Explosive play generation is mediocre (52nd percentile) - lacking big-play ability<br/>
    • Red zone TD rate is poor (39th percentile) - settling for field goals<br/>
    • Recent trends are negative across the board - offense declining down stretch<br/>
    """
    
    elements.append(Paragraph(commentary, styles['Body']))
    
    return elements

def create_defensive_analysis(styles, data):
    """Detailed defensive analysis commentary"""
    elements = []
    
    elements.append(Paragraph("DEFENSIVE ANALYSIS", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 8))
    
    commentary = """
    <b>Ohio State Defense (Elite - 94th Percentile EPA/Play Allowed)</b><br/><br/>
    
    Matt Patricia's defense is dominant across the board. The Buckeyes rank elite in rush defense (94th percentile), 
    limiting explosives (97th percentile - only allowing 4.0% explosive plays), and 3rd down defense (94th percentile). 
    The pass rush generates sacks at an elite rate (97th percentile, 11.0% sack rate).<br/><br/>
    
    <b>Strength:</b> Run defense is a brick wall - 33.2% rush success allowed (95th percentile), -0.151 EPA/rush 
    allowed (94th percentile). Miami's weak run game faces a nightmare matchup.<br/><br/>
    
    <b>Concern:</b> Red zone defense has been vulnerable lately - allowed 41.7% TD rate in last 3 games vs 18.6% 
    season average. Late-game fatigue showing with stuff rate dropping from 25.4% to 18.7% in 2nd half.<br/><br/>
    
    <b>Miami Defense (Very Good - 90th Percentile EPA/Play Allowed)</b><br/><br/>
    
    Miami's defense has been a strength, ranking elite in pass defense (90th percentile EPA/pass allowed) and 
    3rd down defense (92nd percentile). The Hurricanes generate pressure at an 84th percentile rate. The defensive 
    line stuff rate (21.5%) matches Ohio State's, showing competitiveness at the line of scrimmage.<br/><br/>
    
    <b>Critical Weakness - RED ZONE DEFENSE:</b><br/>
    This is Miami's Achilles heel. The Hurricanes rank just 7th percentile in red zone TD rate allowed (43.6%). 
    When opponents get inside the 20, they score touchdowns at an alarming rate. Against Ohio State's 
    97th percentile scoring offense, this is a recipe for disaster.<br/><br/>
    
    <b>Other Concerns:</b><br/>
    • TFL rate is good (13.8% vs 8.9%) but comes with boom-or-bust tendencies<br/>
    • Defense trending wrong direction - EPA/play allowed went from -0.088 to +0.011 in last 3<br/>
    • Rush defense has cratered recently - from 0.010 EPA allowed to 0.191 in last 3 games<br/>
    • D-line gets worn down: YPC allowed jumps from 3.49 (1H) to 4.25 (2H)<br/>
    """
    
    elements.append(Paragraph(commentary, styles['Body']))
    
    return elements

def create_los_commentary(styles, data):
    """Line of scrimmage detailed analysis"""
    elements = []
    
    elements.append(Paragraph("LINE OF SCRIMMAGE BREAKDOWN", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 8))
    
    commentary = """
    <b>The Battle in the Trenches Favors Ohio State</b><br/><br/>
    
    <b>Run Game Progression - Who Controls the 2nd Half?</b><br/><br/>
    
    Good offensive lines "take over" games as they wear down defenses. The data shows:<br/><br/>
    
    <b>Ohio State O-Line:</b> Dominant 2nd half improvement<br/>
    • 1st Half: 42.7% success rate, 4.74 YPC<br/>
    • 2nd Half: 51.3% success rate (+8.6%), 5.23 YPC (+0.49)<br/>
    • Q4 specifically: 5.82 YPC with 51.3% success rate - they close games on the ground<br/><br/>
    
    <b>Miami O-Line:</b> Modest improvement but better at avoiding negative plays<br/>
    • 1st Half: 44.1% success rate, 4.21 YPC<br/>
    • 2nd Half: 47.0% success rate (+2.9%), 4.88 YPC (+0.67)<br/>
    • Stuffed rate (14.3% vs 18.8%) favors Miami - getting stuffed less often<br/><br/>
    
    <b>Defensive Line Durability</b><br/><br/>
    
    <b>Ohio State D-Line:</b> Some late-game fatigue but still dominant<br/>
    • Stuff rate drops from 25.4% (1H) to 18.7% (2H) - fewer plays stopped in backfield<br/>
    • However, YPC allowed stays consistent (3.61 to 3.65) - bend but don't break<br/><br/>
    
    <b>Miami D-Line:</b> Gets worn down as game progresses<br/>
    • Stuff rate drops from 25.5% to 22.9% - fewer stops in backfield<br/>
    • YPC allowed jumps from 3.49 to 4.25 (+0.76) - significant leakage<br/>
    • Against OSU's improving run game, this is a bad combination<br/><br/>
    
    <b>Obvious Situation Performance - Can They Execute When Everyone Knows What's Coming?</b><br/><br/>
    
    <b>Clock Killing (Q4, Leading by 7+):</b><br/>
    • Ohio State: 63% rush rate, 51.2% success, 5.8 YPC - they put games away<br/>
    • Miami: 45% rush rate, 46.4% success, 4.2 YPC - less dominant, more pass-reliant<br/><br/>
    
    <b>Short Yardage Defense (higher stuff rate = better):</b><br/>
    • Ohio State: 38.9% stuff rate, 50.0% success allowed - they win these battles<br/>
    • Miami: 20.0% stuff rate, 73.3% success allowed - they get pushed around<br/><br/>
    
    <b>Pass Rush in Obvious Pass Situations:</b><br/>
    • Ohio State: 16.4% sack rate on 3rd and long - elite pressure<br/>
    • Miami: 3.8% sack rate on 3rd and long - no pressure when needed<br/>
    """
    
    elements.append(Paragraph(commentary, styles['Body']))
    
    return elements

def create_los_discourse_analysis(styles, data):
    """Analysis addressing the discourse about Miami being better up front"""
    elements = []
    
    elements.append(Paragraph("ADDRESSING THE NARRATIVE: IS MIAMI BETTER UP FRONT?", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 8))
    
    discourse = """
    <b>The Popular Narrative vs. The Data</b><br/><br/>
    
    There has been significant discourse suggesting Miami has an advantage in the trenches - that their 
    offensive and defensive lines are superior to Ohio State's. This narrative has been fueled by recruiting 
    rankings, individual player accolades, and the eye test. But what does the statistical evidence actually show?<br/><br/>
    
    <b>OFFENSIVE LINE COMPARISON</b><br/><br/>
    
    <u>Run Blocking Metrics:</u><br/>
    • <b>Yards Per Carry:</b> OSU 5.01 vs Miami 4.63 - <b>Advantage: Ohio State</b><br/>
    • <b>Rush Success Rate:</b> OSU 48.8% vs Miami 47.2% - <b>Advantage: Ohio State</b><br/>
    • <b>Explosive Rush Rate:</b> OSU 15.1% vs Miami 13.1% - <b>Advantage: Ohio State</b><br/>
    • <b>Stuffed Rate:</b> Miami 14.3% vs OSU 18.8% - <b>Advantage: Miami</b> (lower is better for offense)<br/><br/>
    
    <u>Pass Protection Metrics:</u><br/>
    • <b>Sack Rate Allowed:</b> Miami 2.5% vs OSU 2.9% - <b>Slight Advantage: Miami</b><br/>
    • <b>Negative Pass Play Rate:</b> Miami 5.2% vs OSU 4.2% - <b>Advantage: Ohio State</b><br/><br/>
    
    <b>Verdict on O-Line:</b> Miami does have a meaningful advantage in stuff rate (getting stuffed less often - 
    14.3% vs 18.8%) and slightly better sack protection. However, Ohio State's offensive line is statistically 
    superior in the metrics that translate to points - yards per carry, success rate, and explosive play generation. 
    The Buckeyes create more positive plays and bigger gains when they run the ball.<br/><br/>
    
    <b>DEFENSIVE LINE COMPARISON</b><br/><br/>
    
    <u>Run Defense Metrics:</u><br/>
    • <b>YPC Allowed:</b> OSU 3.85 vs Miami 4.16 - <b>Advantage: Ohio State</b><br/>
    • <b>Rush Success Allowed:</b> OSU 33.2% vs Miami 36.5% - <b>Advantage: Ohio State</b><br/>
    • <b>Stuff Rate:</b> Both at 21.5% - <b>Even</b> (higher is better for defense)<br/>
    • <b>TFL Rate:</b> Miami 13.8% vs OSU 8.9% - <b>Advantage: Miami</b><br/><br/>
    
    <u>Pass Rush Metrics:</u><br/>
    • <b>Sack Rate:</b> OSU 8.7% vs Miami 8.1% - <b>Slight Advantage: Ohio State</b><br/>
    • <b>Disruption Rate:</b> OSU 11.0% vs Miami 9.3% - <b>Advantage: Ohio State</b><br/>
    • <b>3rd and Long Sack Rate:</b> OSU 16.4% vs Miami 3.8% - <b>Massive Advantage: Ohio State</b><br/><br/>
    
    <b>Verdict on D-Line:</b> The stuff rates are identical (21.5%), so neither team has an advantage there. 
    While Miami does generate more tackles for loss (13.8% vs 8.9%), Ohio State is superior in the metrics 
    that matter most - YPC allowed, success rate allowed, and pass rush effectiveness in obvious situations. 
    The 16.4% vs 3.8% sack rate on 3rd and long is a staggering disparity.<br/><br/>
    
    <b>THE CRITICAL QUESTION: WHY DOES MIAMI HAVE MORE TFLs BUT WORSE RUN DEFENSE?</b><br/><br/>
    
    This paradox reveals something important about Miami's defensive line play. A higher TFL rate combined 
    with worse overall run defense metrics suggests:<br/><br/>
    
    1. <b>Boom-or-Bust Run Fits:</b> Miami's D-line is aggressive and penetrating, which creates TFLs but also 
    creates gaps. When they win, they win big (TFL). When they lose, they lose big (chunk runs).<br/><br/>
    
    2. <b>Lack of Consistency:</b> Ohio State's defense is more assignment-sound. They may not create as many 
    splash plays, but they rarely give up the explosive runs that Miami surrenders.<br/><br/>
    
    3. <b>Scheme vs. Talent:</b> Miami may have talented individual players who can make plays in the backfield, 
    but the overall defensive front lacks the cohesion and gap discipline of Ohio State's unit.<br/><br/>
    
    <b>2ND HALF PERFORMANCE - THE ULTIMATE TEST</b><br/><br/>
    
    Elite offensive lines impose their will as games progress. Here's what happens in the 2nd half:<br/><br/>
    
    <b>Run Game Success Rate Change (1H to 2H):</b><br/>
    • Ohio State: +8.6% improvement (42.7% to 51.3%)<br/>
    • Miami: +2.9% improvement (44.1% to 47.0%)<br/><br/>
    
    <b>Run Defense Deterioration (1H to 2H):</b><br/>
    • Ohio State YPC Allowed: 3.61 to 3.65 (+0.04) - <b>Holds steady</b><br/>
    • Miami YPC Allowed: 3.49 to 4.25 (+0.76) - <b>Significant leakage</b><br/><br/>
    
    <b>Defensive Stuff Rate Change (1H to 2H):</b><br/>
    • Ohio State: 25.4% to 18.7% (-6.7%) - Some fatigue but still effective<br/>
    • Miami: 25.5% to 22.9% (-2.6%) - Less dropoff but also less impactful overall<br/><br/>
    
    This is the most damning evidence against the "Miami is better up front" narrative. Ohio State's 
    offensive line gets BETTER as games go on while Miami's defensive line gets WORSE. The Buckeyes 
    are built to grind teams down; Miami is not.<br/><br/>
    
    <b>FINAL ASSESSMENT</b><br/><br/>
    
    The statistical evidence does not support the narrative that Miami is better in the trenches. 
    In fact, it suggests the opposite:<br/><br/>
    
    • <b>Offensive Line:</b> Ohio State is clearly superior in run blocking efficiency and explosiveness; 
    Miami has edge only in avoiding negative plays (lower stuff rate 14.3% vs 18.8%)<br/>
    • <b>Defensive Line:</b> Stuff rates are even (21.5%), but Ohio State is superior in run defense 
    and pass rush effectiveness; Miami's higher TFL rate doesn't translate to better overall defense<br/>
    • <b>Durability:</b> Ohio State's lines maintain or improve performance; Miami's deteriorate<br/>
    • <b>Situational:</b> Ohio State dominates in obvious run/pass situations (38.9% vs 20.0% short yardage stuff rate)<br/><br/>
    
    Miami's higher TFL rate and lower offensive stuff rate are the lone bright spots, but they come at 
    the cost of consistency and gap integrity. Ohio State's approach - disciplined, assignment-sound, 
    and relentless - is statistically proven to be more effective.<br/><br/>
    
    <b>Bottom Line:</b> The "Miami is better up front" narrative appears to be based more on 
    perception, recruiting rankings, and individual highlight plays than actual performance data. 
    The numbers tell a clear story: <b>Ohio State controls the line of scrimmage.</b><br/>
    """
    
    elements.append(Paragraph(discourse, styles['Body']))
    
    return elements

def create_keys_to_game(styles, data):
    """Keys to the game for both teams"""
    elements = []
    
    elements.append(Paragraph("KEYS TO THE GAME", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 10))
    
    # Miami Keys
    elements.append(Paragraph("MIAMI MUST:", styles['TeamMiami']))
    elements.append(Spacer(1, 4))
    
    miami_keys = """
    <b>1. Convert in the Red Zone</b><br/>
    Miami's 39th percentile red zone TD rate meets OSU's 70th percentile red zone D. The Hurricanes 
    cannot settle for field goals against an offense that will score touchdowns. Need 70%+ TD rate to compete.<br/><br/>
    
    <b>2. Keep Carson Beck Upright</b><br/>
    OSU's pass rush is elite (16.4% sack rate on 3rd and long). Miami's protection has been decent (2.5% sack 
    rate allowed) but must be perfect. Quick passing game and Beck's pocket presence will be essential.<br/><br/>
    
    <b>3. Limit Explosive Plays</b><br/>
    Miami's 79th percentile explosive rate allowed is good, but OSU's ability to hit big plays when needed 
    is dangerous. One 40+ yard play can swing momentum entirely.<br/><br/>
    
    <b>4. Win 1st Quarter</b><br/>
    Miami performs best tied or with small lead. Their EPA craters when trailing by 7+ (-0.177). 
    A slow start could snowball into a blowout given OSU's 2nd half dominance.<br/><br/>
    
    <b>5. Generate Pressure on Obvious Passing Downs</b><br/>
    Miami's 3.8% sack rate on 3rd and long is a major weakness. The Hurricanes need to find a way to 
    disrupt Julian Sayin's rhythm and get off the field on third down.<br/>
    """
    elements.append(Paragraph(miami_keys, styles['Body']))
    
    elements.append(Spacer(1, 12))
    
    # OSU Keys
    elements.append(Paragraph("OHIO STATE MUST:", styles['TeamOSU']))
    elements.append(Spacer(1, 4))
    
    osu_keys = """
    <b>1. Establish the Run Early</b><br/>
    OSU's run game starts slow (42.7% 1H success) but dominates late (51.3% 2H, 5.82 Q4 YPC). 
    Committing to the run even when it's not working early will pay dividends in the 4th quarter.<br/><br/>
    
    <b>2. Attack Miami's Red Zone Defense</b><br/>
    Miami's 7th percentile red zone TD rate allowed is a glaring weakness. OSU's 61st percentile red zone 
    TD rate is average - they need to punch it in every trip. No field goals inside the 10.<br/><br/>
    
    <b>3. Pressure Carson Beck on 3rd Down</b><br/>
    The Buckeyes' 16.4% sack rate on 3rd and long is elite. Beck has been sacked more in last 3 games (5.3% 
    vs 3.5% season). Bring pressure and force mistakes - Beck's efficiency drops under duress.<br/><br/>
    
    <b>4. Win the Line of Scrimmage Battle Late</b><br/>
    OSU's offensive line gets better as games go on (+8.6% success rate 2H). Miami's D-line fades (allows 
    +0.76 YPC more in 2H). The Buckeyes should pound the rock in the 4th quarter to seal victory.<br/><br/>
    
    <b>5. Control the Clock in the 4th</b><br/>
    OSU is built to grind out wins. Their clock-killing ability (63% rush rate, 5.8 YPC when up 7+) 
    should close this game if they have a lead entering the 4th quarter.<br/>
    """
    elements.append(Paragraph(osu_keys, styles['Body']))
    
    return elements

def create_final_verdict(styles, data):
    """Final analysis and prediction"""
    elements = []
    
    elements.append(Paragraph("FINAL ANALYSIS & VERDICT", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=2, color=HEADER_BG))
    elements.append(Spacer(1, 10))
    
    verdict = """
    <b>THE MATCHUP AT A GLANCE</b><br/><br/>
    
    This game pits Ohio State's historically efficient offense against a Miami team that has been 
    better than expected defensively but has significant structural weaknesses. The statistics paint 
    a clear picture of disparity:<br/><br/>
    
    <b>Where Ohio State Wins (39 statistical advantages):</b><br/>
    • Overall offensive efficiency (91st vs 69th percentile)<br/>
    • Run game efficiency and 2nd half dominance (5.82 Q4 YPC)<br/>
    • 3rd down offense (99th vs 85th percentile)<br/>
    • Red zone defense (70th vs 7th percentile - massive gap)<br/>
    • Limiting explosives on defense (97th vs 79th percentile)<br/>
    • Pass rush in obvious situations (16.4% vs 3.8% sack rate on 3rd and long)<br/><br/>
    
    <b>Where Miami Wins (8 statistical advantages):</b><br/>
    • Pass defense efficiency (90th vs 87th percentile)<br/>
    • Pressure rate generated (84th vs 72nd percentile)<br/>
    • 1st down defense (94th vs 83rd percentile)<br/>
    • Offensive stuff rate (14.3% vs 18.8% - getting stuffed less)<br/>
    • TFL rate (13.8% vs 8.9%)<br/><br/>
    
    <b>X-FACTORS TO WATCH</b><br/><br/>
    
    <b>1. Miami's Red Zone Defense vs OSU's Scoring Machine</b><br/>
    This is THE matchup. Miami allows TDs on 43.6% of red zone trips (7th percentile). Ohio State 
    scores on 58.3% of drives (97th percentile). If OSU reaches the red zone 5+ times, expect 3-4 TDs minimum.<br/><br/>
    
    <b>2. OSU's 4th Quarter Run Game</b><br/>
    The Buckeyes' 5.82 YPC and 51.3% success rate in Q4 is elite. Miami's defense allows 4.25 YPC in the 2nd half 
    (vs 3.49 in 1st half). If this game is within one score entering the 4th, OSU's ability to grind will be decisive.<br/><br/>
    
    <b>3. Can Miami Generate Pressure on Obvious Passing Downs?</b><br/>
    Miami's 3.8% sack rate on 3rd and long is abysmal vs OSU's 16.4%. The Hurricanes need their front to win 
    battles they haven't been winning all season.<br/><br/>
    
    <b>4. 2nd Half Line of Scrimmage Battle</b><br/>
    OSU's O-line improves as games progress (+8.6% success rate 2H). Miami's D-line fades (allows 
    +0.76 YPC more in 2H). This durability gap could be the difference in a close game.<br/><br/>
    
    <b>BOTTOM LINE</b><br/><br/>
    
    Ohio State is the more complete team by a significant margin. They have elite units on both sides 
    of the ball, control the line of scrimmage in obvious situations, and have the statistical profile 
    of a team built for playoff success. Miami has a quality defense and an efficient quarterback in 
    Carson Beck, but the Hurricanes' red zone struggles and inability to generate pressure in key 
    moments are significant handicaps against a team of Ohio State's caliber.<br/><br/>
    
    <b>Miami's Path to Victory:</b> Convert 75%+ red zone TDs, Carson Beck plays mistake-free football, 
    defense generates 4+ sacks to disrupt OSU's rhythm, and limit explosive plays. Even then, 
    it requires OSU's recent offensive regression to continue.<br/><br/>
    
    <b>Ohio State's Path to Victory:</b> Establish run game early, attack red zone aggressively, 
    generate pressure on Beck (especially on 3rd down), and control the 4th quarter. The Buckeyes' 
    statistical advantages are so comprehensive that they simply need to play their game.<br/><br/>
    
    <b>Statistical Lean: OHIO STATE</b> - The gap in overall efficiency, red zone performance, and 
    situational pass rush is too significant. Ohio State's ability to impose their will in the 
    2nd half against a Miami defense that gets worn down (4.25 YPC allowed in 2H vs 3.49 in 1H) 
    is the decisive factor.<br/>
    """
    
    elements.append(Paragraph(verdict, styles['Body']))
    
    return elements

# ============================================================================
# MAIN
# ============================================================================

def generate_report():
    print("Loading data from CSV files...")
    data = load_all_data()
    
    print("Creating PDF document...")
    doc = SimpleDocTemplate(
        "Miami_vs_OhioState_Comprehensive_Report.pdf",
        pagesize=letter,
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch
    )
    
    styles = create_styles()
    elements = []
    
    # Build report sections
    print("Building report sections...")
    
    # Page 1: Title and Summary
    elements.extend(create_title_block(styles, data))
    elements.extend(create_summary_counts(styles, data))
    elements.append(PageBreak())
    
    # Page 2: Executive Summary (NEW COMMENTARY)
    elements.extend(create_executive_summary(styles, data))
    elements.append(PageBreak())
    
    # Page 3-4: Full Comparison
    elements.extend(create_full_comparison(styles, data))
    elements.append(PageBreak())
    
    # Page 5: Head-to-Head Matrix
    elements.extend(create_matchup_matrix(styles, data))
    elements.append(PageBreak())
    
    # Page 6: Edges
    elements.extend(create_matchup_edges(styles, data))
    elements.append(Spacer(1, 10))
    elements.extend(create_strengths_vulnerabilities(styles, data))
    elements.append(PageBreak())
    
    # Page 7: Offensive Analysis (NEW COMMENTARY)
    elements.extend(create_offensive_analysis(styles, data))
    elements.append(PageBreak())
    
    # Page 8: Defensive Analysis (NEW COMMENTARY)
    elements.extend(create_defensive_analysis(styles, data))
    elements.append(PageBreak())
    
    # Page 9: Trends
    elements.extend(create_trend_analysis(styles, data))
    elements.append(PageBreak())
    
    # Page 10: Quarter Analysis
    elements.extend(create_quarter_analysis(styles, data))
    elements.append(Spacer(1, 10))
    elements.extend(create_late_close_analysis(styles, data))
    elements.append(PageBreak())
    
    # Page 11: Situational
    elements.extend(create_situational_analysis(styles, data))
    elements.append(Spacer(1, 10))
    elements.extend(create_field_position_analysis(styles, data))
    elements.append(PageBreak())
    
    # Page 12: Drive & Pressure
    elements.extend(create_drive_efficiency(styles, data))
    elements.append(Spacer(1, 10))
    elements.extend(create_pressure_analysis(styles, data))
    elements.append(PageBreak())
    
    # Page 13: Game Script & Tempo
    elements.extend(create_game_script(styles, data))
    elements.append(Spacer(1, 10))
    elements.extend(create_tempo_analysis(styles, data))
    elements.append(PageBreak())
    
    # Page 14: LINE OF SCRIMMAGE SUMMARY
    elements.extend(create_los_summary(styles, data))
    elements.append(PageBreak())
    
    # Page 15: Run Game Progression
    elements.extend(create_run_progression_analysis(styles, data))
    elements.append(Spacer(1, 10))
    elements.extend(create_obvious_run_analysis(styles, data))
    elements.append(PageBreak())
    
    # Page 16: Run Defense Progression
    elements.extend(create_run_defense_progression(styles, data))
    elements.append(Spacer(1, 10))
    elements.extend(create_run_defense_obvious(styles, data))
    elements.append(PageBreak())
    
    # Page 17: Protection & Pass Rush in Obvious Situations
    elements.extend(create_protection_obvious_analysis(styles, data))
    elements.append(Spacer(1, 10))
    elements.extend(create_pass_rush_obvious_analysis(styles, data))
    elements.append(PageBreak())
    
    # Page 18: LOS Commentary (NEW COMMENTARY)
    elements.extend(create_los_commentary(styles, data))
    elements.append(PageBreak())
    
    # Page 19: LOS Discourse Analysis - Is Miami Better Up Front? (NEW COMMENTARY)
    elements.extend(create_los_discourse_analysis(styles, data))
    elements.append(PageBreak())
    
    # Page 20: Keys to the Game (NEW COMMENTARY)
    elements.extend(create_keys_to_game(styles, data))
    elements.append(PageBreak())
    
    # Page 21: Final Verdict (NEW COMMENTARY)
    elements.extend(create_final_verdict(styles, data))
    
    # Generate PDF
    print("Generating PDF...")
    doc.build(elements)
    
    print("=" * 70)
    print("  COMPREHENSIVE REPORT GENERATED:")
    print("  Miami_vs_OhioState_Comprehensive_Report.pdf")
    print("=" * 70)

if __name__ == "__main__":
    generate_report()
