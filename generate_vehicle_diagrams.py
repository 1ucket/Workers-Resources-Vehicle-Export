import pandas as pd
import matplotlib.pyplot as plt
import os
import argparse
import unicodedata

def clean_text(text):
    return unicodedata.normalize('NFKD', str(text)).encode('ascii', 'ignore').decode('ascii')

def plot_vehicle_availability(excel_file):
    output_dir = "vehicle_diagrams"
    os.makedirs(output_dir, exist_ok=True)
    plt.rcParams['font.family'] = 'Arial'

    xl = pd.ExcelFile(excel_file)
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)

        # Clean up and convert types
        df = df.dropna(subset=['Name', 'StartYear', 'EndYear', 'VehicleType', 'TransportType'])
        df['StartYear'] = pd.to_numeric(df['StartYear'], errors='coerce')
        df['EndYear'] = pd.to_numeric(df['EndYear'], errors='coerce')
        df = df.dropna(subset=['StartYear', 'EndYear'])

        # Group by VehicleType + TransportType
        grouped = df.groupby(['VehicleType', 'TransportType'])

        for (vtype, ttype), group in grouped:
            group_sorted = group.sort_values(by='StartYear', ascending=True)

            # Clamp end year to 2020
            group_sorted['EndYear'] = group_sorted['EndYear'].clip(upper=2020)

            fig, ax = plt.subplots(figsize=(10, 0.5 * len(group_sorted)))
            vehicle_names = []
            for idx, row in enumerate(group_sorted.itertuples()):
                vehicle_label = clean_text(row.Name)
                vehicle_names.append(vehicle_label)
                # Get capacity or skill if available (adjust keys as needed)
                capacity = getattr(row, 'Capacity', None)
                skill = getattr(row, 'Skill', None)
                label_value = capacity if capacity and capacity != 'N/A' else skill

                ax.barh(
                    y=vehicle_label,
                    width=row.EndYear - row.StartYear,
                    left=row.StartYear,
                    height=0.5,
                    align='center'
                )
                if label_value and label_value != 'N/A':
                    ax.text(
                        row.StartYear + (row.EndYear - row.StartYear) / 2,
                        vehicle_label,
                        f"{label_value}",
                        va='center',
                        ha='center',
                        color='white',
                        fontsize=8,
            fontweight='bold'
        )

            # Check for gaps or boundary adjacency between vehicles in time
            # Create a timeline of (start, end) sorted by start year
            # After sorting and clipping EndYear:
            # After sorting and clipping EndYear:
            intervals = sorted([(row.StartYear, row.EndYear) for row in group_sorted.itertuples()])

            # Merge only strictly overlapping intervals (not touching)
            merged_intervals = []
            for start, end in intervals:
                if not merged_intervals:
                    merged_intervals.append([start, end])
                else:
                    last_start, last_end = merged_intervals[-1]
                    # Merge only if intervals strictly overlap (start < last_end)
                    if start < last_end:
                        merged_intervals[-1][1] = max(last_end, end)
                    else:
                        merged_intervals.append([start, end])

            # Find gaps between merged intervals (including touching gaps)
            for i in range(len(merged_intervals) - 1):
                gap_start = merged_intervals[i][1]
                gap_end = merged_intervals[i+1][0]
                if gap_end >= gap_start:  # include touching gaps
                    print(f"⚠️ Gap detected in {vtype.title()} - {ttype.title()} between years {int(gap_start)} and {int(gap_end)}")
                    # Shade the gap area
                    ax.axvspan(gap_start, gap_end, color='red', alpha=0.2)
                    # Label the gap in the middle of the shaded area, below the bars
                    mid_gap = (gap_start + gap_end) / 2
                    ax.text(mid_gap, -1, "Gap", color='red', fontsize=8, rotation=90, va='bottom', ha='center')




            ax.set_xlim(group_sorted['StartYear'].min(), 2020)
            ax.set_xlabel("Year")
            ax.set_ylabel("Vehicle")
            ax.set_yticks(range(len(vehicle_names)))
            ax.set_yticklabels(vehicle_names)
            ax.set_title(f"{vtype.title()} - {ttype.title()}")
            ax.grid(True, axis='x', linestyle='--', alpha=0.5)

            plt.tight_layout()
            filename = f"{vtype}_{ttype}.png".replace(" ", "_")
            plt.savefig(os.path.join(output_dir, filename))
            plt.close()

    print(f"✅ Diagrams saved to: {output_dir}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate vehicle availability diagrams from Excel export")
    parser.add_argument(
        "excel_file",
        nargs="?",
        default="vehicles.xlsx",
        help="Path to the Excel file (default: vehicles.xlsx in current folder)"
    )
    args = parser.parse_args()

    plot_vehicle_availability(args.excel_file)
