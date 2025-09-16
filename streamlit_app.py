import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import geopandas as gpd
import math
from io import BytesIO
import re

# Custom CSS for the dashboard
st.markdown("""
<style>
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        max-width: none;
    }
    
    h1 {
        color: #2c3e50;
        text-align: center;
        font-weight: 700;
        margin-bottom: 2rem;
    }
    
    h2, h3 {
        color: #34495e;
        border-bottom: 2px solid #3498db;
        padding-bottom: 0.5rem;
    }
    
    .stButton > button {
        background: linear-gradient(45deg, #3498db, #2980b9);
        color: white;
        border: none;
        border-radius: 25px;
        padding: 0.5rem 2rem;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

def create_chiefdom_mapping():
    """Create mapping between GPS data chiefdom names and shapefile FIRST_CHIE names"""
    chiefdom_mapping = {
        # BO District mappings
        "Bo City": "BO TOWN",
        "Badjia": "BADJIA",
        "Bargbo": "BAGBO",
        "Bagbwe": "BAGBWE(BAGBE)",
        "Baoma": "BOAMA",
        "Bongor": "BONGOR",
        "Bumpeh": "BUMPE NGAO",
        "Gbo": "GBO",
        "Jaiama": "JAIAMA",
        "Kakua": "KAKUA",
        "Komboya": "KOMBOYA",
        "Lugbu": "LUGBU",
        "Niawa Lenga": "NIAWA LENGA",
        "Selenga": "SELENGA",
        "Tinkoko": "TIKONKO",
        "Valunia": "VALUNIA",
        "Wonde": "WONDE",
        
        # BOMBALI District mappings
        "Biriwa": "BIRIWA",
        "Bombali Sebora": "BOMBALI SEBORA",
        "Bombali Serry": "BOMBALI SIARI",
        "Gbanti (Bombali)": "GBANTI",
        "Gbanti": "GBANTI",
        "Gbendembu": "GBENDEMBU",
        "Kamaranka": "KAMARANKA",
        "Magbaimba Ndohahun": "MAGBAIMBA NDORWAHUN",
        "Makarie": "MAKARI",
        "Mara": "MARA",
        "Ngowahun": "N'GOWAHUN",
        "Paki Masabong": "PAKI MASABONG",
        "Safroko Limba": "SAFROKO LIMBA",
        "Makeni City": "MAKENI CITY",
    }
    return chiefdom_mapping

def map_chiefdom_name(chiefdom_name, mapping):
    """Map chiefdom name from GPS data to shapefile name"""
    if pd.isna(chiefdom_name):
        return None
    
    chiefdom_name = str(chiefdom_name).strip()
    
    # Direct match
    if chiefdom_name in mapping:
        return mapping[chiefdom_name]
    
    # Case-insensitive match
    for key, value in mapping.items():
        if key.upper() == chiefdom_name.upper():
            return value
    
    # Partial match (contains)
    for key, value in mapping.items():
        if key.upper() in chiefdom_name.upper() or chiefdom_name.upper() in key.upper():
            return value
    
    # Return original if no mapping found
    return chiefdom_name

def calculate_itn_totals_per_row(df, row_idx):
    """Calculate ITN totals for a specific row - CONSISTENT METHOD"""
    boys_total = 0
    girls_total = 0
    left_total = 0
    
    # Calculate boys and girls ITNs across all classes
    for class_num in range(1, 6):
        boys_col = f"How many boys in Class {class_num} received ITNs?"
        girls_col = f"How many girls in Class {class_num} received ITNs?"
        
        if boys_col in df.columns:
            boys_val = df[boys_col].iloc[row_idx]
            if pd.notna(boys_val):
                boys_total += int(boys_val)
        
        if girls_col in df.columns:
            girls_val = df[girls_col].iloc[row_idx]
            if pd.notna(girls_val):
                girls_total += int(girls_val)
    
    # ITNs left at school (single value per school) - FIXED: Only counted once per school
    left_col = "ITNs left at the school for pupils who were absent."
    if left_col in df.columns:
        left_val = df[left_col].iloc[row_idx]
        if pd.notna(left_val):
            left_total = int(left_val)
    
    # Total ITNs = Boys + Girls + Left at School
    total_distributed = boys_total + girls_total + left_total
    
    return {
        'boys': boys_total,
        'girls': girls_total, 
        'left': left_total,
        'total_distributed': total_distributed
    }

def calculate_itn_totals_for_dataframe(df):
    """Calculate ITN totals for entire dataframe - CONSISTENT METHOD with detailed breakdown"""
    total_boys = 0
    total_girls = 0
    total_left = 0
    
    # Calculate boys and girls ITNs across all classes with detailed tracking
    boys_by_class = {}
    girls_by_class = {}
    
    for class_num in range(1, 6):
        boys_col = f"How many boys in Class {class_num} received ITNs?"
        girls_col = f"How many girls in Class {class_num} received ITNs?"
        
        if boys_col in df.columns:
            class_boys = int(df[boys_col].fillna(0).sum())
            boys_by_class[f"Class {class_num}"] = class_boys
            total_boys += class_boys
        
        if girls_col in df.columns:
            class_girls = int(df[girls_col].fillna(0).sum())
            girls_by_class[f"Class {class_num}"] = class_girls
            total_girls += class_girls
    
    # FIXED: ITNs left at school - this is per school, not per class
    # Only sum this once across all schools, not multiplied by classes
    left_col = "ITNs left at the school for pupils who were absent."
    if left_col in df.columns:
        total_left = int(df[left_col].fillna(0).sum())
    
    # Total ITNs = Boys + Girls + Left at School
    total_distributed = total_boys + total_girls + total_left
    
    return {
        'boys': total_boys,
        'girls': total_girls,
        'left': total_left,
        'total_distributed': total_distributed,
        'boys_by_class': boys_by_class,
        'girls_by_class': girls_by_class
    }

def extract_itn_data_from_excel(df):
    """Extract ITN coverage data from the Excel file using CONSISTENT calculation"""
    # Create empty lists to store extracted data
    districts, chiefdoms, total_enrollment, distributed_itns = [], [], [], []
    
    # Get chiefdom mapping
    chiefdom_mapping = create_chiefdom_mapping()
    
    # Process each row in the "Scan QR code" column
    for idx, qr_text in enumerate(df["Scan QR code"]):
        if pd.isna(qr_text):
            districts.append(None)
            chiefdoms.append(None)
            total_enrollment.append(0)
            distributed_itns.append(0)
            continue
            
        # Extract values using regex patterns
        district_match = re.search(r"District:\s*([^\n]+)", str(qr_text))
        district = district_match.group(1).strip() if district_match else None
        districts.append(district)
        
        chiefdom_match = re.search(r"Chiefdom:\s*([^\n]+)", str(qr_text))
        original_chiefdom = chiefdom_match.group(1).strip() if chiefdom_match else None
        
        # Map chiefdom name to match shapefile
        mapped_chiefdom = map_chiefdom_name(original_chiefdom, chiefdom_mapping)
        chiefdoms.append(mapped_chiefdom)
        
        # Calculate total enrollment (sum of all class enrollments)
        enrollment_total = 0
        for class_num in range(1, 6):  # Classes 1-5
            enrollment_col = f"How many pupils are enrolled in Class {class_num}?"
            if enrollment_col in df.columns:
                class_enrollment = df[enrollment_col].iloc[idx]
                if pd.notna(class_enrollment):
                    enrollment_total += int(class_enrollment)
        
        total_enrollment.append(enrollment_total)
        
        # FIXED: Calculate ITN totals using consistent method
        itn_data = calculate_itn_totals_per_row(df, idx)
        distributed_itns.append(itn_data['total_distributed'])
    
    # Create a new DataFrame with extracted values
    itn_df = pd.DataFrame({
        "District": districts,
        "Chiefdom": chiefdoms,
        "Total_Enrollment": total_enrollment,
        "Distributed_ITNs": distributed_itns
    })
    
    return itn_df

def generate_summaries(df):
    """Generate District, Chiefdom, and Gender summaries - CONSISTENT WITH FIRST DOCUMENT"""
    summaries = {}
    
    # Overall Summary
    overall_summary = {
        'total_schools': len(df),
        'total_districts': len(df['District'].dropna().unique()) if 'District' in df.columns else 0,
        'total_chiefdoms': len(df['Chiefdom'].dropna().unique()) if 'Chiefdom' in df.columns else 0,
        'total_boys': 0,
        'total_girls': 0,
        'total_enrollment': 0,
        'total_itn': 0,
        'total_left': 0
    }
    
    # Calculate totals using the correct columns
    for class_num in range(1, 6):
        # Total enrollment from "Number of enrollments in class X"
        enrollment_col = f"How many pupils are enrolled in Class {class_num}?"
        if enrollment_col in df.columns:
            overall_summary['total_enrollment'] += int(df[enrollment_col].fillna(0).sum())
        
        # Boys and girls for gender analysis AND ITN calculation
        boys_col = f"How many boys in Class {class_num} received ITNs?"
        girls_col = f"How many girls in Class {class_num} received ITNs?"
        
        if boys_col in df.columns:
            overall_summary['total_boys'] += int(df[boys_col].fillna(0).sum())
        if girls_col in df.columns:
            overall_summary['total_girls'] += int(df[girls_col].fillna(0).sum())
    
    # FIXED: ITNs left at school - this should be per school, not per class
    itn_left_col = "ITNs left at the school for pupils who were absent."
    if itn_left_col in df.columns:
        overall_summary['total_left'] += int(df[itn_left_col].fillna(0).sum())
    
    # Total ITNs = boys + girls + left (all ITNs distributed or allocated)
    overall_summary['total_itn'] = overall_summary['total_boys'] + overall_summary['total_girls'] + overall_summary['total_left']
    
    # Calculate coverage
    overall_summary['coverage'] = (overall_summary['total_itn'] / overall_summary['total_enrollment'] * 100) if overall_summary['total_enrollment'] > 0 else 0
    overall_summary['itn_remaining'] = overall_summary['total_enrollment'] - overall_summary['total_itn']
    
    summaries['overall'] = overall_summary
    
    # District Summary
    district_summary = []
    for district in df['District'].dropna().unique():
        district_data = df[df['District'] == district]
        district_stats = {
            'district': district,
            'schools': len(district_data),
            'chiefdoms': len(district_data['Chiefdom'].dropna().unique()),
            'boys': 0,
            'girls': 0,
            'enrollment': 0,
            'itn': 0,
            'left': 0
        }
        
        for class_num in range(1, 6):
            # Total enrollment from "Number of enrollments in class X"
            enrollment_col = f"How many pupils are enrolled in Class {class_num}?"
            if enrollment_col in district_data.columns:
                district_stats['enrollment'] += int(district_data[enrollment_col].fillna(0).sum())
            
            # Boys and girls for gender analysis AND ITN calculation
            boys_col = f"How many boys in Class {class_num} received ITNs?"
            girls_col = f"How many girls in Class {class_num} received ITNs?"
            
            if boys_col in district_data.columns:
                district_stats['boys'] += int(district_data[boys_col].fillna(0).sum())
            if girls_col in district_data.columns:
                district_stats['girls'] += int(district_data[girls_col].fillna(0).sum())
        
        # ITNs left at school - this should be per school, not per class
        itn_left_col = "ITNs left at the school for pupils who were absent."
        if itn_left_col in district_data.columns:
            district_stats['left'] += int(district_data[itn_left_col].fillna(0).sum())
        
        # Total ITNs = boys + girls + left (all ITNs distributed or allocated)
        district_stats['itn'] = district_stats['boys'] + district_stats['girls'] + district_stats['left']
        
        # Calculate coverage
        district_stats['coverage'] = (district_stats['itn'] / district_stats['enrollment'] * 100) if district_stats['enrollment'] > 0 else 0
        district_stats['itn_remaining'] = district_stats['enrollment'] - district_stats['itn']
        
        district_summary.append(district_stats)
    
    summaries['district'] = district_summary
    
    # Chiefdom Summary - UPDATED TO INCLUDE ITNs LEFT
    chiefdom_summary = []
    for district in df['District'].dropna().unique():
        district_data = df[df['District'] == district]
        for chiefdom in district_data['Chiefdom'].dropna().unique():
            chiefdom_data = district_data[district_data['Chiefdom'] == chiefdom]
            chiefdom_stats = {
                'district': district,
                'chiefdom': chiefdom,
                'schools': len(chiefdom_data),
                'boys': 0,
                'girls': 0,
                'enrollment': 0,
                'itn': 0,
                'left': 0
            }
            
            for class_num in range(1, 6):
                # Total enrollment from "Number of enrollments in class X"
                enrollment_col = f"How many pupils are enrolled in Class {class_num}?"
                if enrollment_col in chiefdom_data.columns:
                    chiefdom_stats['enrollment'] += int(chiefdom_data[enrollment_col].fillna(0).sum())
                
                # Boys and girls for gender analysis AND ITN calculation
                boys_col = f"How many boys in Class {class_num} received ITNs?"
                girls_col = f"How many girls in Class {class_num} received ITNs?"
                
                if boys_col in chiefdom_data.columns:
                    chiefdom_stats['boys'] += int(chiefdom_data[boys_col].fillna(0).sum())
                if girls_col in chiefdom_data.columns:
                    chiefdom_stats['girls'] += int(chiefdom_data[girls_col].fillna(0).sum())
            
            # ITNs left at school - this should be per school, not per class
            itn_left_col = "ITNs left at the school for pupils who were absent."
            if itn_left_col in chiefdom_data.columns:
                chiefdom_stats['left'] += int(chiefdom_data[itn_left_col].fillna(0).sum())
            
            # Total ITNs = boys + girls + left (all ITNs distributed or allocated)
            chiefdom_stats['itn'] = chiefdom_stats['boys'] + chiefdom_stats['girls'] + chiefdom_stats['left']
            
            # Calculate coverage
            chiefdom_stats['coverage'] = (chiefdom_stats['itn'] / chiefdom_stats['enrollment'] * 100) if chiefdom_stats['enrollment'] > 0 else 0
            chiefdom_stats['itn_remaining'] = chiefdom_stats['enrollment'] - chiefdom_stats['itn']
            
            chiefdom_summary.append(chiefdom_stats)
    
    summaries['chiefdom'] = chiefdom_summary
    
    return summaries

def get_coverage_color(coverage_percent):
    """Get color based on coverage percentage"""
    if coverage_percent < 20:
        return '#d32f2f'  # Red
    elif coverage_percent < 40:
        return '#f57c00'  # Orange
    elif coverage_percent < 60:
        return '#fbc02d'  # Yellow
    elif coverage_percent < 80:
        return '#388e3c'  # Light Green
    elif coverage_percent < 100:
        return '#1976d2'  # Blue
    else:
        return '#4a148c'  # Purple (100% coverage)

def create_itn_coverage_dashboard(gdf, itn_df, district_name, cols=4):
    """Create ITN coverage dashboard optimized for Word document export - FIXED CALCULATIONS"""
    
    # Filter shapefile for the district
    district_gdf = gdf[gdf['FIRST_DNAM'] == district_name].copy()
    
    if len(district_gdf) == 0:
        st.error(f"No chiefdoms found for {district_name} district in shapefile")
        return None
    
    # Get unique chiefdoms from shapefile
    chiefdoms = sorted(district_gdf['FIRST_CHIE'].dropna().unique())
    
    # Calculate rows needed
    rows = math.ceil(len(chiefdoms) / cols)
    
    # Optimize figure size for Word document (16:10 aspect ratio works well)
    fig_width = 16  # Width for Word document
    fig_height = rows * 3.5  # Height per row optimized for Word
    
    # Create subplot figure optimized for Word export
    fig, axes = plt.subplots(rows, cols, figsize=(fig_width, fig_height))
    fig.suptitle(f'{district_name} District - ITN Coverage Analysis', 
                 fontsize=18, fontweight='bold', y=0.98)
    
    # Ensure axes is always 2D array
    if rows == 1:
        axes = axes.reshape(1, -1)
    elif cols == 1:
        axes = axes.reshape(-1, 1)
    
    # Plot each chiefdom
    for idx, chiefdom in enumerate(chiefdoms):
        row = idx // cols
        col = idx % cols
        ax = axes[row, col]
        
        # Filter shapefile for this specific chiefdom
        chiefdom_gdf = district_gdf[district_gdf['FIRST_CHIE'] == chiefdom].copy()
        
        # Filter ITN data for this district and chiefdom
        district_data = itn_df[itn_df["District"].str.upper() == district_name.upper()].copy()
        chiefdom_data = district_data[district_data["Chiefdom"] == chiefdom].copy()
        
        # FIXED: Calculate totals for this chiefdom using consistent method
        enrollment_total = int(chiefdom_data["Total_Enrollment"].sum()) if len(chiefdom_data) > 0 else 0
        itns_total = int(chiefdom_data["Distributed_ITNs"].sum()) if len(chiefdom_data) > 0 else 0
        
        # Calculate coverage percentage
        coverage_percent = (itns_total / enrollment_total * 100) if enrollment_total > 0 else 0
        coverage_percent = min(coverage_percent, 100)  # Cap at 100%
        
        # Get color based on coverage
        coverage_color = get_coverage_color(coverage_percent)
        
        # Plot chiefdom boundary with coverage color
        chiefdom_gdf.plot(ax=ax, color=coverage_color, edgecolor='black', alpha=0.8, linewidth=1.5)
        
        # Create ITN coverage text (n, m) format
        itn_text = f"({itns_total}, {enrollment_total})"
        
        # Set title with ITN coverage information (optimized font size for Word)
        ax.set_title(f'{chiefdom}\n{itn_text}', 
                    fontsize=10, fontweight='bold', pad=8)
        
        # Add coverage percentage in the center of the chiefdom
        if len(chiefdom_gdf) > 0:
            # Get center of chiefdom
            bounds = chiefdom_gdf.total_bounds
            center_x = (bounds[0] + bounds[2]) / 2
            center_y = (bounds[1] + bounds[3]) / 2
            
            # Add coverage percentage text in the center
            ax.text(center_x, center_y, f"{coverage_percent:.0f}%", 
                   fontsize=14, fontweight='bold', color='white', 
                   ha='center', va='center',
                   bbox=dict(boxstyle='round,pad=0.5', facecolor='black', alpha=0.7))
        
        # Remove axis labels and ticks for cleaner look
        ax.set_xticks([])
        ax.set_yticks([])
        ax.set_xlabel('')
        ax.set_ylabel('')
        
        # Remove the box frame
        for spine in ax.spines.values():
            spine.set_visible(False)
        
        # Add very light grid
        ax.grid(True, alpha=0.2, linestyle='--', linewidth=0.5)
        
        # Set equal aspect ratio
        ax.set_aspect('equal')
        
        # Set bounds to chiefdom extent with minimal padding for better fit
        bounds = chiefdom_gdf.total_bounds
        padding = 0.005  # Reduced padding for better fit in Word
        ax.set_xlim(bounds[0] - padding, bounds[2] + padding)
        ax.set_ylim(bounds[1] - padding, bounds[3] + padding)
    
    # Hide empty subplots
    total_plots = rows * cols
    for idx in range(len(chiefdoms), total_plots):
        row = idx // cols
        col = idx % cols
        axes[row, col].set_visible(False)
    
    # Optimize layout for Word document
    plt.tight_layout()
    plt.subplots_adjust(top=0.90, hspace=0.35, wspace=0.25)  # Increased space below title
    
    return fig

def generate_simple_summary(itn_df):
    """Generate simple summary with just totals and coverage - USING CONSISTENT CALCULATIONS"""
    
    summary_data = []
    
    # District totals
    for district in ["BO", "BOMBALI"]:
        district_data = itn_df[itn_df["District"].str.upper() == district.upper()]
        
        # FIXED: Using consistent calculation method
        total_enrollment = int(district_data["Total_Enrollment"].sum())
        total_itns_distributed = int(district_data["Distributed_ITNs"].sum())
        coverage = (total_itns_distributed / total_enrollment * 100) if total_enrollment > 0 else 0
        
        summary_data.append({
            'Level': 'District',
            'Name': district,
            'Total_Enrollment': total_enrollment,
            'ITNs_Distributed': total_itns_distributed,
            'Coverage': f"{coverage:.1f}%"
        })
    
    # Chiefdom totals for each district
    for district in ["BO", "BOMBALI"]:
        district_data = itn_df[itn_df["District"].str.upper() == district.upper()]
        
        # Group by chiefdom
        chiefdoms = district_data['Chiefdom'].dropna().unique()
        
        for chiefdom in sorted(chiefdoms):
            chiefdom_data = district_data[district_data['Chiefdom'] == chiefdom]
            
            # FIXED: Using consistent calculation method
            total_enrollment = int(chiefdom_data["Total_Enrollment"].sum())
            total_itns_distributed = int(chiefdom_data["Distributed_ITNs"].sum())
            coverage = (total_itns_distributed / total_enrollment * 100) if total_enrollment > 0 else 0
            
            summary_data.append({
                'Level': f'{district} Chiefdom',
                'Name': chiefdom,
                'Total_Enrollment': total_enrollment,
                'ITNs_Distributed': total_itns_distributed,
                'Coverage': f"{coverage:.1f}%"
            })
    
    return summary_data

# Streamlit App
st.title("🛡️ Section 3: ITN Coverage Analysis")
st.markdown("**ITN distribution effectiveness by chiefdom**")

# Color legend
st.markdown("""
### Coverage Color Legend:
- 🔴 **Red**: < 20%
- 🟠 **Orange**: 20-39%
- 🟡 **Yellow**: 40-59%
- 🟢 **Light Green**: 60-79%
- 🔵 **Blue**: 80-99%
- 🟣 **Purple**: 100%+
""")

# Coverage format explanation
st.markdown("""
### ITN Coverage Format:
- **Title**: `(ITNs Distributed, Total Enrollment)`
- **Center**: Coverage percentage
- **Colors**: Same as above color legend
- **ITN Calculation**: Boys ITNs + Girls ITNs + ITNs Left at School for Absent Pupils
""")

# File Information
st.info("""
**📁 Embedded Files:** `SBD_Final_data_dissemination_7_15_2025.xlsx` | `Chiefdom2021.shp`  
**📊 Layout:** Fixed 4-column grid optimized for Word export
""")

# Load the embedded data files
try:
    # Load Excel file (embedded)
    df_original = pd.read_excel("SBD_Final_data_dissemination_pmi_evolve_2025_09.xlsx")
    st.success(f"✅ Excel file loaded successfully! Found {len(df_original)} records.")
    
except Exception as e:
    st.error(f"❌ Error loading Excel file: {e}")
    st.info("💡 Make sure 'SBD_Final_data_dissemination_7_15_2025.xlsx' is in the same directory as this app")
    st.stop()

# Load shapefile (embedded)
try:
    gdf = gpd.read_file("Chiefdom2021.shp")
    st.success(f"✅ Shapefile loaded successfully! Found {len(gdf)} features.")
    
except Exception as e:
    st.error(f"❌ Could not load shapefile: {e}")
    st.info("💡 Make sure 'Chiefdom2021.shp' and supporting files (.dbf, .shx, .prj) are in the same directory as this app")
    st.stop()

# Dashboard Settings - Fixed configuration
columns = 4  # Fixed to 4 columns for optimal Word export
show_itn_details = True  # Always show ITN data details

# Extract ITN coverage data
try:
    itn_df = extract_itn_data_from_excel(df_original)
    st.success(f"✅ ITN data extracted successfully! Found {len(itn_df)} records.")
    
    # Debug: Show comprehensive ITNs calculation breakdown
    with st.expander("🔍 Debug: ITN Calculation Verification"):
        st.write("**Available columns containing 'ITN' or 'left':**")
        itn_columns = [col for col in df_original.columns if 'itn' in col.lower() or 'left' in col.lower()]
        for col in itn_columns:
            st.write(f"- {col}")
        
        # Calculate using consistent method
        st.write("**📊 ITN Distribution Breakdown (Using Consistent Method):**")
        total_data = calculate_itn_totals_for_dataframe(df_original)
        
        # Show boys by class
        st.write("**Boys ITNs by Class:**")
        for class_name, count in total_data['boys_by_class'].items():
            st.write(f"  - {class_name}: {count:,}")
        st.write(f"  - **Total Boys**: {total_data['boys']:,}")
        
        # Show girls by class  
        st.write("**Girls ITNs by Class:**")
        for class_name, count in total_data['girls_by_class'].items():
            st.write(f"  - {class_name}: {count:,}")
        st.write(f"  - **Total Girls**: {total_data['girls']:,}")
        
        # Show left at school details
        st.write("**ITNs Left at School Analysis:**")
        left_col = "ITNs left at the school for pupils who were absent."
        if left_col in df_original.columns:
            left_values = df_original[left_col].dropna()
            st.write(f"  - Number of schools with 'left' data: {len(left_values)}")
            st.write(f"  - Min left at school: {left_values.min()}")
            st.write(f"  - Max left at school: {left_values.max()}")
            st.write(f"  - Mean left at school: {left_values.mean():.1f}")
            st.write(f"  - **Total Left at School**: {total_data['left']:,}")
            
            # Show sample of left values
            st.write("  - Sample 'left' values from first 10 schools:")
            for i in range(min(10, len(left_values))):
                st.write(f"    School {i+1}: {int(left_values.iloc[i])}")
        else:
            st.error(f"  - Column '{left_col}' not found!")
        
        st.write(f"**🎯 FINAL TOTALS:**")
        st.write(f"- Boys ITNs: {total_data['boys']:,}")
        st.write(f"- Girls ITNs: {total_data['girls']:,}")
        st.write(f"- Left at School: {total_data['left']:,}")
        st.write(f"- **Grand Total ITNs: {total_data['total_distributed']:,}**")
        
        calculated_total = int(itn_df["Distributed_ITNs"].sum())
        st.write(f"- **DataFrame Total: {calculated_total:,}**")
        
        if calculated_total != total_data['total_distributed']:
            st.error(f"❌ MISMATCH! Expected {total_data['total_distributed']:,} but got {calculated_total:,}")
            st.write("**Investigating the mismatch...**")
            
            # Check if we're double-counting or missing something
            st.write("**Manual verification from raw data:**")
            manual_boys = sum([df_original[f"How many boys in Class {i} received ITNs?"].fillna(0).sum() for i in range(1, 6) if f"How many boys in Class {i} received ITNs?" in df_original.columns])
            manual_girls = sum([df_original[f"How many girls in Class {i} received ITNs?"].fillna(0).sum() for i in range(1, 6) if f"How many girls in Class {i} received ITNs?" in df_original.columns])
            manual_left = df_original[left_col].fillna(0).sum() if left_col in df_original.columns else 0
            manual_total = manual_boys + manual_girls + manual_left
            
            st.write(f"Manual calculation: Boys={manual_boys:,}, Girls={manual_girls:,}, Left={manual_left:,}, Total={manual_total:,}")
            
        else:
            st.success("✅ ITN calculations match perfectly!")
            
        # What should the correct total be?
        st.write("**🤔 Expected ITN Total Check:**")
        st.write("Based on your feedback, the total should NOT be 216,851.")
        st.write(f"Current calculation gives: {total_data['total_distributed']:,}")
        st.write("Please verify if this breakdown looks correct:")
        st.write(f"- Boys receiving ITNs: {total_data['boys']:,}")
        st.write(f"- Girls receiving ITNs: {total_data['girls']:,}")  
        st.write(f"- ITNs left for absent pupils: {total_data['left']:,}")
        st.write("")
        st.write("**Question: Should we exclude the 'ITNs left at school' from the total?**")
        st.write("If yes, the total would be:", f"{total_data['boys'] + total_data['girls']:,}")
            
            
except Exception as e:
    st.error(f"Error extracting ITN data: {e}")
    itn_df = pd.DataFrame()

if show_itn_details and len(itn_df) > 0:
    # Display ITN data information
    st.subheader("📊 ITN Data Overview")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_enrollment = int(itn_df["Total_Enrollment"].sum())
        st.metric("Total Enrollment", f"{total_enrollment:,}")
    
    with col2:
        total_distributed = int(itn_df["Distributed_ITNs"].sum())
        st.metric("ITNs Distributed", f"{total_distributed:,}")
    
    with col3:
        overall_coverage = (total_distributed / total_enrollment * 100) if total_enrollment > 0 else 0
        st.metric("Overall Coverage", f"{overall_coverage:.1f}%")
    
    with col4:
        schools_with_distribution = len(itn_df[itn_df["Distributed_ITNs"] > 0])
        st.metric("Schools with ITNs", f"{schools_with_distribution:,}")

# Create dashboards
st.header("🛡️ ITN Coverage Dashboards")

# BO District ITN Coverage Dashboard
st.subheader("BO District - ITN Coverage")

with st.spinner("Generating BO District ITN coverage dashboard..."):
    try:
        fig_bo_itn = create_itn_coverage_dashboard(gdf, itn_df, "BO", columns)
        if fig_bo_itn:
            st.pyplot(fig_bo_itn)
            
            # Save figure option
            buffer_bo_itn = BytesIO()
            fig_bo_itn.savefig(buffer_bo_itn, format='png', dpi=300, bbox_inches='tight')
            buffer_bo_itn.seek(0)
            
            st.download_button(
                label="📥 Download BO District ITN Coverage Dashboard (PNG)",
                data=buffer_bo_itn,
                file_name="BO_District_ITN_Coverage_Dashboard.png",
                mime="image/png"
            )
        else:
            st.warning("Could not generate BO District ITN coverage dashboard")
    except Exception as e:
        st.error(f"Error generating BO District ITN coverage dashboard: {e}")

st.divider()

# BOMBALI District ITN Coverage Dashboard
st.subheader("BOMBALI District - ITN Coverage")

with st.spinner("Generating BOMBALI District ITN coverage dashboard..."):
    try:
        fig_bombali_itn = create_itn_coverage_dashboard(gdf, itn_df, "BOMBALI", columns)
        if fig_bombali_itn:
            st.pyplot(fig_bombali_itn)
            
            # Save figure option
            buffer_bombali_itn = BytesIO()
            fig_bombali_itn.savefig(buffer_bombali_itn, format='png', dpi=300, bbox_inches='tight')
            buffer_bombali_itn.seek(0)
            
            st.download_button(
                label="📥 Download BOMBALI District ITN Coverage Dashboard (PNG)",
                data=buffer_bombali_itn,
                file_name="BOMBALI_District_ITN_Coverage_Dashboard.png",
                mime="image/png"
            )
        else:
            st.warning("Could not generate BOMBALI District ITN coverage dashboard")
    except Exception as e:
        st.error(f"Error generating BOMBALI District ITN coverage dashboard: {e}")

# ITN Analysis
st.header("📈 ITN Distribution Analysis")

if len(itn_df) > 0:
    # Calculate ITN statistics using consistent method
    bo_data = itn_df[itn_df["District"].str.upper() == "BO"]
    bombali_data = itn_df[itn_df["District"].str.upper() == "BOMBALI"]
    
    # ITN metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        bo_enrollment = int(bo_data["Total_Enrollment"].sum())
        bo_distributed = int(bo_data["Distributed_ITNs"].sum())
        bo_coverage = (bo_distributed / bo_enrollment * 100) if bo_enrollment > 0 else 0
        st.metric("BO District ITN Coverage", f"{bo_coverage:.1f}%", f"{bo_distributed}/{bo_enrollment}")
    
    with col2:
        bombali_enrollment = int(bombali_data["Total_Enrollment"].sum())
        bombali_distributed = int(bombali_data["Distributed_ITNs"].sum())
        bombali_coverage = (bombali_distributed / bombali_enrollment * 100) if bombali_enrollment > 0 else 0
        st.metric("BOMBALI District ITN Coverage", f"{bombali_coverage:.1f}%", f"{bombali_distributed}/{bombali_enrollment}")
    
    with col3:
        total_enrollment = int(itn_df["Total_Enrollment"].sum())
        total_distributed = int(itn_df["Distributed_ITNs"].sum())
        overall_coverage = (total_distributed / total_enrollment * 100) if total_enrollment > 0 else 0
        st.metric("Overall ITN Coverage", f"{overall_coverage:.1f}%", f"{total_distributed}/{total_enrollment}")
    
    with col4:
        # Calculate chiefdoms with good ITN coverage (>= 60%)
        good_itn_coverage_count = 0
        total_chiefdoms = 0
        
        for district in ["BO", "BOMBALI"]:
            district_data = itn_df[itn_df["District"].str.upper() == district.upper()]
            chiefdoms = district_data['Chiefdom'].dropna().unique()
            
            for chiefdom in chiefdoms:
                chiefdom_data = district_data[district_data['Chiefdom'] == chiefdom]
                enrollment = int(chiefdom_data["Total_Enrollment"].sum())
                distributed = int(chiefdom_data["Distributed_ITNs"].sum())
                coverage = (distributed / enrollment * 100) if enrollment > 0 else 0
                
                if coverage >= 60:
                    good_itn_coverage_count += 1
                total_chiefdoms += 1
        
        good_itn_coverage_percent = (good_itn_coverage_count / total_chiefdoms * 100) if total_chiefdoms > 0 else 0
        st.metric("Chiefdoms with Good ITN Coverage", f"{good_itn_coverage_percent:.0f}%", f"{good_itn_coverage_count}/{total_chiefdoms}")
    
    # Detailed ITN coverage table
    st.subheader("📋 Detailed ITN Coverage by Chiefdom")
    
    detailed_itn_coverage = []
    for district in ["BO", "BOMBALI"]:
        district_data = itn_df[itn_df["District"].str.upper() == district.upper()]
        chiefdoms = sorted(district_data['Chiefdom'].dropna().unique())
        
        for chiefdom in chiefdoms:
            chiefdom_data = district_data[district_data['Chiefdom'] == chiefdom]
            enrollment = int(chiefdom_data["Total_Enrollment"].sum())
            distributed = int(chiefdom_data["Distributed_ITNs"].sum())
            coverage = (distributed / enrollment * 100) if enrollment > 0 else 0
            
            # Determine status
            if coverage >= 80:
                status = "✅ Excellent"
            elif coverage >= 60:
                status = "🟢 Good"
            elif coverage >= 40:
                status = "🟡 Fair"
            elif coverage >= 20:
                status = "🟠 Poor"
            else:
                status = "🔴 Critical"
            
            detailed_itn_coverage.append({
                'District': district,
                'Chiefdom': chiefdom,
                'Total Enrollment': enrollment,
                'ITNs Distributed': distributed,
                'Coverage %': f"{coverage:.1f}%",
                'Status': status
            })
    
    itn_coverage_df = pd.DataFrame(detailed_itn_coverage)
    st.dataframe(itn_coverage_df, use_container_width=True)
    
    # Distribution Summary
    st.subheader("📊 Distribution Summary")
    
    try:
        summary_data = generate_simple_summary(itn_df)
        summary_df = pd.DataFrame(summary_data)
        
        st.dataframe(summary_df, use_container_width=True)
    
    except Exception as e:
        st.error(f"Error generating distribution summary: {e}")

# Show the calculation formula clearly
st.header("🧮 ITN Calculation Formula")
st.info("""
**Complete ITN Distribution Formula:**

ITNs Distributed = Boys ITNs (Classes 1-5) + Girls ITNs (Classes 1-5) + ITNs Left at School for Absent Pupils

**Components:**
- **Boys ITNs**: Sum of boys who received ITNs across all classes (1-5)
- **Girls ITNs**: Sum of girls who received ITNs across all classes (1-5)  
- **ITNs Left at School**: ITNs reserved at school for pupils who were absent during distribution

This ensures complete accounting of all ITN distribution efforts including direct distribution and reserves for absent students.
""")

# Memory optimization - close matplotlib figures
plt.close('all')

# Footer
st.markdown("---")
st.markdown("**🛡️ Section 3: ITN Coverage Analysis | School-Based Distribution Analysis**")
