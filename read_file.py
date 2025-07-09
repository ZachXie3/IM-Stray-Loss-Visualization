import pandas as pd
from pathlib import Path
import logging

# Set up logging
log_dir = Path("logs")
log_dir.mkdir(exist_ok=True)
logging.basicConfig(
    filename=log_dir / "process_log.txt",
    filemode="w",
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

def process_design_files(folder="Design"):
    folder_path = Path(folder)
    excel_files = list(folder_path.glob("*.xls")) + list(folder_path.glob("*.xlsx"))

    missing_sheet_files = []
    no_correlation_files = []
    extracted_results = []

    for file in excel_files:
        try:
            engine = "openpyxl" if file.suffix == ".xlsx" else "xlrd"

            xl = pd.ExcelFile(file, engine=engine)
            if "DesignSheet" not in xl.sheet_names:
                missing_sheet_files.append(file.name)
                continue

            df = xl.parse(sheet_name="DesignSheet", header=None)

            h36_value = str(df.iat[35, 7])
            if " / " not in h36_value:
                no_correlation_files.append(file.name)
                continue

            result = extract_design_data(df)
            result["filename"] = file.name
            extracted_results.append(result)

        except Exception as e:
            logging.error(f"Error processing {file.name}: {e}")

    # Save logs
    if missing_sheet_files:
        logging.warning("Files missing 'DesignSheet':")
        for f in missing_sheet_files:
            logging.warning(f"- {f}")

    if no_correlation_files:
        logging.warning("Files with invalid H36 (no ' / '):")
        for f in no_correlation_files:
            logging.warning(f"- {f}")

    return extracted_results


def extract_design_data(df: pd.DataFrame) -> dict:
    try:
        data = {
            "pole": df.iat[4, 0],
            "hp": df.iat[4, 1],
            "rpm": df.iat[4, 2],
            "voltage": df.iat[4, 4],
            "frequency": df.iat[4, 5],
            "frame": df.iat[4, 9],
            "airgap": df.iat[7, 1],
            "core": df.iat[9, 1],
            "stator_od": df.iat[10, 1],
            "stator_id": df.iat[11, 1],
            "stator_core_depth": df.iat[12, 1],
            "rotor_od": df.iat[10, 2],
            "rotor_id": df.iat[11, 2],
            "rotor_core_depth": df.iat[12, 2],
            "stator_slot_num": df.iat[15, 1],
            "rotor_slot_num": df.iat[15, 2],
            "stator_slot_shape": df.iat[16, 1],
            "rotor_slot_shape": df.iat[16, 2],
            "stator_slot_depth": df.iat[18, 1],
            "rotor_slot_depth": df.iat[18, 2],
            "stator_slot_width": df.iat[20, 1],
            "stator_tip_depth": df.iat[21, 1],
            "rotor_tip_depth": df.iat[21, 2],
            "low_cage_depth": df.iat[54, 3],
            "low_cage_upper_width": df.iat[52, 3],
            "low_cage_lower_width": df.iat[53, 3],
            "mid_tooth_width": df.iat[53, 5],
            "stray_loss": df.iat[37, 9],
            "wdg_temp_rise": df.iat[34, 4],            
            "flux_pri_core": df.iat[29, 6],
            "flux_pri_tooth": df.iat[30, 6],
            "flux_sec_core": df.iat[31, 6],
            "flux_sec_tooth": df.iat[32, 6],
            "flux_air_gap": df.iat[33, 6],
        }

        def extract_loss(val):
            try:
                if isinstance(val, str) and " / " in val:
                    return float(val.split(" / ")[1].strip())
                return float(val)
            except Exception:
                return 0.0  # fallback to 0.0 to avoid type issues

        data["no_load_loss"] = extract_loss(df.iat[34, 7])
        data["pri_loss"] = extract_loss(df.iat[35, 7])
        data["sec_loss"] = extract_loss(df.iat[36, 7])
        data["core_loss"] = extract_loss(df.iat[36, 9])
        data["fri_wdg_loss"] = extract_loss(df.iat[37, 7])

        pri = float(data.get("pri_loss", 0) or 0)
        sec = float(data.get("sec_loss", 0) or 0)
        core = float(data.get("core_loss", 0) or 0)
        fri = float(data.get("fri_wdg_loss", 0) or 0)
        stray = float(data.get("stray_loss", 0) or 0)


        total_loss = pri + sec + core + fri + stray
        data["total_loss"] = total_loss

        if total_loss > 0:
            data["pri_loss_pct"] = pri / total_loss
            data["sec_loss_pct"] = sec / total_loss
            data["core_loss_pct"] = core / total_loss
            data["fri_wdg_loss_pct"] = fri / total_loss
            data["stray_loss_pct"] = stray / total_loss
        else:
            data["pri_loss_pct"] = None
            data["sec_loss_pct"] = None
            data["core_loss_pct"] = None
            data["fri_wdg_loss_pct"] = None
            data["stray_loss_pct"] = None

        try:
            rotor_core_depth = float(data["rotor_core_depth"])
            rotor_slot_depth = float(data["rotor_slot_depth"])
            mid_tooth_width = float(data["mid_tooth_width"])

            data["rotor_slot_core_ratio"] = (
                rotor_slot_depth / rotor_core_depth if rotor_core_depth else None
            )
            data["rotor_tooth_depth_width_ratio"] = (
                rotor_slot_depth / mid_tooth_width if mid_tooth_width else None
            )
            data["pri_loss_stray_loss_ratio"] = (
                pri / stray if stray else None
            )
            data["core_loss_stray_loss_ratio"] = (
                core / stray if stray else None
            )
            data["sec_loss_stray_loss_ratio"] = (
                sec / stray if stray else None
            )

        except Exception:
            data["rotor_slot_core_ratio"] = None
            data["rotor_tooth_depth_width_ratio"] = None

        return data

    except Exception as e:
        logging.error(f"Error in extract_design_data: {e}")
        return {}

if __name__ == "__main__":
    results = process_design_files("Design")
    df_results = pd.DataFrame(results)
    output_file = Path("results") / "design_results.xlsx"
    output_file.parent.mkdir(exist_ok=True)
    df_results.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")
