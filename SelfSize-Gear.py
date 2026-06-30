# ===========================================
# 1: SYSTEM IMPORTS AND GLOBAL CONFIGURATIONS
# ===========================================

import json
import keras
import math
import matplotlib.pyplot as plt
import numpy as np
import os
import pandas as pd
import seaborn as sns
import sys
import tensorflow as tf
import win32com.client

# ---Global Visual Styling ---
# For graph consistency.
sns.set_theme (style="whitegrid")
plt.rcParams['font.family']='sans-serif'

# ==========================================
# 2. EXPLORATORY DATA ANALYSIS & BACKFILLING
# ==========================================

# 1. Alloy Classification (Carburizing vs. Through-Hardening)

def assign_alloy_classes(df_working):

    # 1. Creates working copy to avoid modifying the source data
    deep_copy_df = df_working.copy(deep=True)

    # 2. Defines boolean arrays with strict parentheses grouping
    is_case_hardening = (deep_copy_df['C_Max'] <= 0.25)
    is_through_hardening = (deep_copy_df['C_Max'] > 0.25) & (deep_copy_df['C_Max'] <= 0.60)
    is_bearing = (deep_copy_df['C_Max'] > 0.60)

    # 3. Applies category labels using localized conditional indexing
    conditions = [is_case_hardening, is_through_hardening, is_bearing,]
    choices = ['Carburizing/Case-Hardening','Through-Hardening/QT','High-Carbon/Bearing',]
    deep_copy_df['Application_Class'] = np.select(conditions, choices, default='Unknown')

    # 4. Returns the modified container
    return deep_copy_df

def plot_metallurgical_metrics(df:pd.DataFrame, x_col: str, y_metric: str, alloy_class="Application_Class", output_dir: str="plots") -> str:

    # 1. Creates the output directory if it doesn't exist.
    os.makedirs(output_dir, exist_ok=True)

    # 2. Dictionary to map column substrings to metallurgical units.
    data = {"Hardness_HB": "HB", "Yield_strength": "MPa", "_Min": "%", "_Max": "%", "UTS_Min": "Mpa", "Allowable_Contact_Mpa": "Mpa"}

    # 3. Dynamically extracts units in one line.
    x_unit = next((f"({unit})"
                   for key, unit in data.items()
                   if key in x_col),
                   "",)
    
    y_unit = next((f"({unit})"
                   for key, unit in data.items()
                   if key in y_metric),
                   "",)

    # 4. Generates the plot
    plt.figure(figsize=(10,6))
    sns.scatterplot(data=df, x=x_col, y=y_metric, hue="Application_Class", s=60, alpha=0.75)
    plt.title(f"{y_metric.title()} vs {x_col.title()} ({alloy_class})")
    plt.xlabel(f"{x_col}{x_unit}")
    plt.ylabel(f"{y_metric}{y_unit}")
    plt.legend(frameon=True, shadow=True, loc="best")

    # 6. Builds the dynamic filename (sanitized to lowercase without spaces)
    clean_metric = y_metric.lower().replace(" ","_")
    clean_class = alloy_class.lower().replace(" ", "_")
    filename = f"scatter_{clean_metric}_{clean_class}.png"

    # 7. Combines directory and filename
    full_path = os.path.join(output_dir, filename)

    # 8. Saves the file
    plt.savefig(full_path, dpi=300, bbox_inches="tight")

    # 9. Clears and closes to free memory
    plt.clf()
    plt.close("all")

    return full_path
    
def load_process_data(config) -> pd.DataFrame:

    # 1. Loads data using configuration.
    df_raw = pd.read_csv(config["filename"], delimiter=config["delimiter"])

    # 2. Processes and captures output into a new variable.
    df_processed = assign_alloy_classes(df_raw)
    return df_processed

def audit_dataset_integrity(df_copy):

    # 1. Working copy created to avoid modifying the source data.
    df_temporary = df_copy.copy(deep=True)

    # 2. Boolean arrays defined with strict grouping.
    uts_expected = df_temporary['Hardness_HB'] * 3.45
    relative_deviation = (df_temporary['UTS_Min'] - uts_expected).abs() / uts_expected
    anomalous = (relative_deviation > 0.15) & (df_temporary['UTS_Min'].notna())
    consistent = (relative_deviation <= 0.15) & (df_temporary['UTS_Min'].notna())
    unverified = df_temporary['UTS_Min'].isna()

    # 3. Category labels applied using localized conditional indexing.
    conditions = [anomalous, consistent, unverified]
    choices = ['Anomalous-Pair','Consistent-Pair','Unverified']
    df_temporary['Mechanical_Class'] = np.select(conditions, choices, default='Unknown')

    # 4. Returns the modified container.
    return df_temporary

def verify_dataset_integrity(df:pd.DataFrame):

    # 1. Critical structural parameters that CANNOT have zero or missing variables. Classification column isolated.
    critical_mechanical_columns = ["Yield_strength", "Hardness_HB", "Allowable_Bending_Mpa", "Allowable_Contact_Mpa"]
    critical_metallurgical_columns = ["C_Max"]
    target_column = df["Application_Class"]
    audit_column = df['Mechanical_Class']

    # 2. Categorical flags and structural empty states globally audited. Classification column checked.
    combined_list = critical_mechanical_columns + critical_metallurgical_columns
    critical_list = [col for col in combined_list if col in df.columns]
    global_invalid_masks = ((df[critical_list].isna()) | (df[critical_list]=='unknown') | (df[critical_list]=='Unknown')).sum().sum()
    invalid_class_mask = ((target_column == 'Unknown') | (target_column.isna())).sum()
    anomaly_class_mask = (audit_column=="Anomalous-Pair").sum()

    # 3. Physically impossible numerical zeros only in targeted design metrics audited.
    all_critical_columns = critical_mechanical_columns + critical_metallurgical_columns
    existing_columns = [col for col in all_critical_columns if col in df.columns]
    zero_metric_errors = (df[existing_columns]==0).sum().sum()

    # 4. Aggregates and notifies without logging valid 0.0% chemistry profiles.
    total_true_issues = global_invalid_masks + zero_metric_errors + invalid_class_mask + anomaly_class_mask

    if (total_true_issues) > 0:
        print(f"⚠️ Verified data leaks and/or empty metrics and/or anomalies detected!")
        return "Failure"
    else:
        print(f"✅ Verified classification and physical consistency!")
        return "Success"

def segregate_data(df_copy, clean_path, quarantine_path):

    # 1. Working copy created to avoid modifying the source data.
    df_temporary = df_copy.copy(deep=True)

    # 2. Structural integrity masks created.
    critical_checklist = ["Yield_strength", "Hardness_HB", "Allowable_Bending_Mpa", "Allowable_Contact_Mpa"]
    empty_cells = (df_temporary[critical_checklist].isin(["unknown", "Unknown"])) | (df_temporary[critical_checklist].isna())
    missing_final = empty_cells.any(axis=1)

    # 3. Physical anomaly mask constructed
    mech_class = df_temporary["Mechanical_Class"]
    mechanical_anomaly_mask = (mech_class=="Anomalous-Pair")

    # 4. Master condition developed to merge the above trackers.
    row_failure_condition = (missing_final) | (mechanical_anomaly_mask)

    # 5. The subsets are extracted and isolated.
    quarantine_subset = df_temporary[row_failure_condition].copy()
    clean_subset = df_temporary[~row_failure_condition].copy()
    quarantine_subset.to_csv(quarantine_path, sep=";", index=False)
    clean_subset.to_csv(clean_path, sep=";", index=False)

    # 6. Isolated containers outputted.
    return clean_subset

# ==========================================
# 3. FEATURE ENGINEERING AND SCALE ALIGNMENT
# ==========================================

# 1. Min-Max Feature Scaling

def scale_engineering_features(df_copy, target_columns_list):

    # 1. Working copy created to avoid modifying source data.
    df_temporary = df_copy.copy(deep=True)

    # 2. Columns are dynamically processed to calculate boundaries.    
    abs_df = df_temporary[target_columns_list].abs()
    current_min = abs_df.min()
    current_max = abs_df.max()

    # 3. Scaling transformation applied (Min-Max Scaling).
    scaling_manifest = {col: (current_min[col], current_max[col])for col in target_columns_list}
    df_temporary[target_columns_list] = (df_temporary[target_columns_list] - current_min) / (current_max - current_min)

    # 4. Returns the results.
    return df_temporary, scaling_manifest

# 2. Physics-Informed Loss Functions Established.

def calculate_physics_informed_penalties(predictions_matrix, scaling_manifest):

    # 1. Manifest used to unscale the neural network outputs back into real-world units.
    y_min, y_max = scaling_manifest["Yield_strength"]
    actual_yield_Mpa = (predictions_matrix["Yield_strength"] * (y_max - y_min) + y_min)

    hb_min, hb_max = scaling_manifest["Hardness_HB"]
    actual_hardness_hb = (predictions_matrix["Hardness_HB"] * (hb_max - hb_min) + hb_min)

    ac_min, ac_max = scaling_manifest["Allowable_Contact_Mpa"]
    actual_contact_pressure = (predictions_matrix["Allowable_Contact_Mpa"] * (ac_max - ac_min) + ac_min)

    # 2. Penalty values initialized.
    empirical_strength_penalty = 0.0
    physical_boundary_penalty = 0.0

    # 3. Metallurgical Law 1: The Hardness-to-Tensile Baseline.
    # The Ultimate Tensile Strength is roughly 3.45 times Brinell Hardness.
    # Yield strength must structurally sit below this UTS value.
    expected_upper_bound = actual_hardness_hb * 3.45
    violations = actual_yield_Mpa - expected_upper_bound
    clamped_violations = np.maximum(violations,0)
    empirical_strength_penalty = np.sum(clamped_violations ** 2)
    
    # 4. Mechanical Law 2: Contact Stress Structural Limits.
    # Contact stress capacity scales in proportion to surface core support.
    physical_limit = 3 * actual_yield_Mpa
    breach_difference = np.maximum(actual_contact_pressure - physical_limit,0)
    physical_boundary_penalty = np.sum(breach_difference ** 2)

    # 5. Total structural deviations compiled and outputted.
    total_physics_loss = empirical_strength_penalty + physical_boundary_penalty
    return total_physics_loss

def train_pinn_sizing_core(df_scaled, scaling_manifest, iterations=100):
    # 1. Inputs from targets are isolated using scaled data frame.
    input_features_to_scale = ("C_Max","Hardness_HB")
    target_features_to_scale = ("Yield_strength","Hardness_HB","Allowable_Contact_Mpa","Allowable_Bending_Mpa")
    input_arrays = tf.convert_to_tensor(df_scaled[list(input_features_to_scale)].values, dtype=tf.float32)
    target_arrays = tf.convert_to_tensor(df_scaled[list(target_features_to_scale)].values, dtype=tf.float32)

    # 2. Iterative neural network training cycle executed.
    pinn_model = tf.keras.Sequential([
        tf.keras.Input(shape=(len(input_features_to_scale),)),
        tf.keras.layers.Dense(64, activation='tanh'),
        tf.keras.layers.Dense(len(target_features_to_scale))
    ])
    optimizer = tf.keras.optimizers.Adam(learning_rate=0.001)

    # 3. Physically optimized model asset saved.
    for _ in range(iterations):
        with tf.GradientTape() as tape:
            raw_outputs = pinn_model(input_arrays, training=True)
            raw_network_predictions = {
                "Yield_strength": raw_outputs[:,0],
                "Hardness_HB": raw_outputs[:,1],
                "Allowable_Contact_Mpa": raw_outputs[:,2],
                "Allowable_Bending_Mpa": raw_outputs[:,3]
            }

            data_driven_loss = tf.reduce_mean(tf.square(target_arrays - raw_outputs))
            total_physics_loss = calculate_physics_informed_penalties(raw_network_predictions, scaling_manifest)
            ultimate_pinn_loss = data_driven_loss + total_physics_loss
        
        grads = tape.gradient(ultimate_pinn_loss, pinn_model.trainable_variables)
        optimizer.apply_gradients(zip(grads, pinn_model.trainable_variables))
    
    # 4. Storage path and saved asset constructed.
    script_directory = os.path.dirname(__file__)
    model_file_name = "gear_sizing_core.keras"
    absolute_model_path = os.path.join(script_directory, "Data", model_file_name)
    pinn_model.save(absolute_model_path)
    print(f"✅ The PINN model is compiled, locked, and saved to disk as 'gear_sizing_core.keras'!")
    return pinn_model

def predict_and_unscale_properties (pinn_model, raw_input_data, scaling_manifest):

    # 1. Incoming user metrics extracted and scaled.
    raw_carbon, raw_hardness = raw_input_data[["C_Max","Hardness_HB"]].values.T
    c_min, c_max = scaling_manifest["C_Max"]
    scaled_carbon = (raw_carbon - c_min)/(c_max - c_min)
    h_min, h_max = scaling_manifest["Hardness_HB"]
    scaled_hardness = (raw_hardness - h_min)/(h_max - h_min)

    # 2. Features for model evaluation formatted.
    input_tensor = tf.convert_to_tensor([[scaled_carbon, scaled_hardness]], dtype=tf.float32)

    # 3. Predictions from the model asset requested.
    scaled_predictions = pinn_model(input_tensor).numpy()[0]

    # 4 & 5. Manifest keys for reversing the scaling transformation used and translated properties outputted.
    ys_min, ys_max = scaling_manifest["Yield_strength"]
    hb_min, hb_max = scaling_manifest["Hardness_HB"]
    ac_min, ac_max = scaling_manifest["Allowable_Contact_Mpa"]
    ab_min, ab_max = scaling_manifest["Allowable_Bending_Mpa"]

    physical_predictions = {
        "Yield_strength": scaled_predictions[0] * (ys_max - ys_min) + ys_min,
        "Hardness_HB": scaled_predictions[1] * (hb_max - hb_min) + hb_min,
        "Allowable_Contact_Mpa": scaled_predictions[2] * (ac_max - ac_min) + ac_min,
        "Allowable_Bending_Mpa": scaled_predictions[3] * (ab_max - ab_min) + ab_min
    }

    return physical_predictions

# ===============================
# 4. AUTODESK CAD API INTEGRATION
# ===============================

# 1. CAD Parameter Payload Compilation

def export_cad_parameter_payload(physical_predictions, gear_geometry_inputs):

    # 1. Absolute file pathways established.
    script_directory = os.path.dirname(__file__)
    export_file_name = "active_gear_parameters.json"
    absolute_export_path = os.path.join(script_directory, "Data", export_file_name)

    # 2. A comprehensive parameter manifest built.
    cad_payload = {}

    # 3. Material performance constraints derived from the injected PINN.
    cad_payload["material_limits"] = {
        "yield_strength_mpa": float(physical_predictions["Yield_strength"]),
        "allowable_contact_mpa": float(physical_predictions["Allowable_Contact_Mpa"]),
        "allowable_bending_mpa": float(physical_predictions["Allowable_Bending_Mpa"])
    }

    # 4. Structural geometry targets are calculated from injected limits.
    # Extract calculated PINN limits and input parameters for cleaner math
    sigma_b = cad_payload["material_limits"]["allowable_bending_mpa"]
    sigma_c = cad_payload["material_limits"]["allowable_contact_mpa"]

    # Safely extract mechanical user inputs (with default fallbacks if missing)
    torque = float(gear_geometry_inputs.get("torque_nm", 100.0))
    pinion_teeth = float(gear_geometry_inputs.get("pinion_teeth", 20.0))
    lewis_y = float(gear_geometry_inputs.get("lew_factor_y", 0.3))

    # --- MECHANICAL EQUATIONS ---
    # A) Target Hardness: Pulled directly from your physical predictions dictionary
    target_hb = float(physical_predictions["Hardness_HB"])

    # B) Module (m): Derived from contact stress limits (simplification of AGMA surface durability)
    # Higher contact limits allow a smaller module; higher torque forces a larger module
    calculated_module = (torque/(sigma_c * pinion_teeth)) ** (1/3) * 10

    # C) Face Width (b): Derived from Lewis Bending Equation (b = Torque / (σ_b * m * Y * radius))
    # Higher bending limit allows a narrower, lighter gear face width
    pitch_radius = (calculated_module * pinion_teeth)/2
    calculated_face_width = torque / (sigma_b * calculated_module * lewis_y * pitch_radius) * 1000

    cad_payload["gear_dimensions"] = {
        "target_hardness_hb":target_hb,
        "face_width_mm":round(float(calculated_face_width),2),
        "module_value":round(float(calculated_module),2)
    }
    # 5. Clean payload file committed to disk.
    # Ensure the "Data" subfolder actually exists before trying to write to it
    os.makedirs(os.path.dirname(absolute_export_path), exist_ok=True)

    # Open a file writer stream at the absolute location
    with open(absolute_export_path, "w", encoding="utf-8") as file_stream:
        # Convert the dictionary structure into a standardized, scannable text format (JSON)
        # 'indent=4' satisfies the "scannable text" condition for human review
        json.dump(cad_payload, file_stream, indent=4)
    
    # The 'with' context manager automatically closes the file stream here 
    print(f"✅ CAD Parameter Bridge File locked and ready for Autodesk ingestion.")
    return absolute_export_path

def generate_3d_gear_geometry():

    # 1. The file path to match the pipeline's output location is established.
    script_directory = os.path.dirname(__file__)
    export_file_name = "active_gear_parameters.json"
    payload_file_path = os.path.join(script_directory, "Data", export_file_name)

    # 2. The data exchange payload is ingested.
    with open(payload_file_path, "r", encoding="utf-8") as file_stream:
        active_parameters = json.load(file_stream)
    # 3. The parameters for geometric construction is extracted.
    local_module = active_parameters["gear_dimensions"]["module_value"]
    local_face_width = active_parameters["gear_dimensions"]["face_width_mm"]
    material_note = active_parameters["gear_dimensions"]["target_hardness_hb"]

    #The user inputs matching the original generation parameters are set.
    local_teeth_count = 20
    pressure_angle = 20.0 # Standard industrial gear pressure angle.

    # 4. The secondary driven dimensions for the CAD model are calculated.
    pitch_diameter = local_module * local_teeth_count
    addendum_value = local_module
    dedendum_value = local_module * 1.25
    outer_diameter = pitch_diameter + (2 * addendum_value)

    # Driven radii for plotting (converted directly to cm for Inventor's database).
    r_outer = (outer_diameter / 2.0) / 10.0
    r_pitch = (pitch_diameter / 2.0) / 10.0
    r_root = ((pitch_diameter - (2 * dedendum_value)) / 2.0) / 10.0
    
    # 5. The CAD modeling canvas is initialized.
    try:
        # The active CAD application framework (Inventor engine) is accessed.
        inventor_app = win32com.client.GetActiveObject("Inventor.Application")

        # A new, blank Part Document asset is created (kPartDocumentObject enum = 12290)
        kPartDocumentObject = 12290
        part_doc = inventor_app.Documents.Add(kPartDocumentObject, "", True)
        part_comp_def = part_doc.ComponentDefinition

        # The custom iProperties metadata container is accessed.
        custom_prop_set = part_doc.PropertySets.Item("Inventor User Defined Properties")

        # The textual metadata attribute named "Required_Hardness" is attached.
        try:
            # The property value is updated if it already exists in the template.
            custom_prop_set.Item("Required_Hardness").Value = material_note
        except Exception:
            # A fresh custom property is created if it does not exist.
            custom_prop_set.Add(material_note, "Required_Hardness")
        
        print(f"Canvas initialized successfully. 'Required_Hardness' set to {material_note}.")
    
    except Exception as e:
        print(f"Failed to initialize canvas in Step 5: {e}", file=sys.stderr)
    
        return False
    
    # 6. 2D Sketching operations executed.
    # A fresh 2D drawing sketch plane on the X-Y baseline is opened (Item 3)
    xy_plane = part_comp_def.WorkPlanes.Item(3)
    active_sketch = part_comp_def.Sketches.Add(xy_plane)
    active_sketch.Name = "Gear_Profile_Sketch"

    trans_geom = inventor_app.TransientGeometry

    # Pitch spacing angle calculated (in radians for Python math)
    spacing_angle = 360.0 / local_teeth_count
    angle_per_tooth = math.radians(spacing_angle)

    # The width of a single tooth at the pitch is approximated (roughly half the pitch space)
    half_tooth_angle = math.pi / (2 * local_teeth_count)

    print("Pre-calculating watertight gear profile coordinate matrix...")
    # Step A: Compile all coordinates sequentially to map out the entire gear perimeter path
    profile_points = []
    for i in range(local_teeth_count):
        # The current tooth's Centerline vector
        center_theta = i * angle_per_tooth
        # The pitch circle's left and right flank angular positions.
        theta_left = center_theta - half_tooth_angle
        theta_right = center_theta + half_tooth_angle

        # Left Flank Root (Start of Tooth)
        x_root_l = r_root * math.cos(theta_left)
        y_root_l = r_root * math.sin(theta_left)
        
        # Left Flank Tip (Outer Land Start)
        x_outer_l = r_outer * math.cos(theta_left + 0.02)
        y_outer_l = r_outer * math.sin(theta_left + 0.02)
        
        # Right Flank Tip (Outer Land End)
        x_outer_r = r_outer * math.cos(theta_right + 0.02)
        y_outer_r = r_outer * math.sin(theta_right + 0.02)
        
        # Right Flank Root (End of Tooth)
        x_root_r = r_root * math.cos(theta_right)
        y_root_r = r_root * math.sin(theta_right)

        # Append the point objects in a precise sequence tracing the continuous outer shape
        profile_points.extend([
            trans_geom.CreatePoint2d(x_root_l, y_root_l),
            trans_geom.CreatePoint2d(x_outer_l, y_outer_l),
            trans_geom.CreatePoint2d(x_outer_r, y_outer_r),
            trans_geom.CreatePoint2d(x_root_r, y_root_r)
        ])

    print("Constructing continuous 2D tooth profile paths via COM point chaining...")
    # Step B: Draw the very first line segment to initialize the chain
    first_line = active_sketch.SketchLines.AddByTwoPoints(profile_points[0], profile_points[1])
    current_end_point = first_line.EndSketchPoint

    # Step C: Chain the rest of the lines by feeding the previous line's actual EndSketchPoint object pointer
    for m in range(2, len(profile_points)):
        next_line = active_sketch.SketchLines.AddByTwoPoints(current_end_point, profile_points[m])
        current_end_point = next_line.EndSketchPoint

    # Step D: Securely connect the final line segment back to the first line's StartSketchPoint to lock the loop
    active_sketch.SketchLines.AddByTwoPoints(current_end_point, first_line.StartSketchPoint)
    
    active_sketch.DeferUpdates = False
    active_sketch.Solve()

    # The solid profile loops are captured cleanly with no open boundary exceptions.
    loop_profiles = active_sketch.Profiles.AddForSolid()

    # 7. 3D Solid feature modeling executed.
    print("Solid extrusion operation initialized...")
    extrude_features = part_comp_def.Features.ExtrudeFeatures

    # The face width parameter is converted from mm to internal cm
    face_width_cm = local_face_width / 10.0

    # Extrusion configurations:
    # kJoinOperation = 20481 (Creates a new solid volume)
    # kSymmetricExtentDirection = 20994 (Symmetric thickness distribution)
    kJoinOperation = 20481
    kSymmetricExtentDirection = 20994

    # The extrusion behavior is initialized and configured.
    extrusion = extrude_features.AddByDistanceExtent(
        loop_profiles,
        face_width_cm,
        kSymmetricExtentDirection,
        kJoinOperation
    )
    print("3D Extrusion executed successfully.")

    # 8. Component is rendered and locked to disk.
    # The viewport camera view orientation is updated to standard Isometric (kIsometricViewOrientation = 10762)
    kIsometricViewOrientation = 10762
    active_view = inventor_app.ActiveView
    active_view.Camera.ViewOrientationType = kIsometricViewOrientation
    active_view.Camera.Fit()
    active_view.Update()

    # absolute_cad_save_path targeting project directory defined.
    # Standard Inventor part extension is .ipt (using base file string name)
    save_directory = os.path.join(script_directory, "Output")
    if not os.path.exists(save_directory):
        os.makedirs(save_directory)
    
    absolute_cad_save_path = os.path.join(save_directory, "pinn_optimized_geometry.ipt")

    # The solid model is permanently saved to disk.
    part_doc.SaveAs(absolute_cad_save_path, False)

    print(f"✅ 3D Solid Model dynamically generated from PINN parameters.")
    return True

if __name__ == "__main__":

    # 1. Loads data directly from previous steps.
    script_directory = os.path.dirname(__file__)
    csv_file_name = "self_size_gear.csv"
    absolute_csv_path = os.path.join(script_directory, "Data", csv_file_name)
    absolute_clean_path = os.path.join(script_directory, "Data", "clean_production_gear.csv")
    absolute_quarantine_path = os.path.join(script_directory, "Data", "quarantine_anomalies.csv")

    prod_config = {"filename":absolute_csv_path, "delimiter":";"}
    df_processed = load_process_data(prod_config)

    # 2. Targets generated column and metrics
    df_processed = assign_alloy_classes(df_processed)
    df_processed = audit_dataset_integrity(df_processed)

    df_current_class = "Application_Class"
    process_metrics = ["Yield_strength","Allowable_Contact_Mpa"]
    critical_cols_metrics = ["Yield_strength","Hardness_HB","Allowable_Bending_Mpa","Allowable_Contact_Mpa","C_Max"]

    # 3. Extracts value counts from the newly-created classification column.
    print(df_processed["Application_Class"].value_counts().to_string(header=False))
    print(df_processed["Mechanical_Class"].value_counts().to_string(header=False))

    # 4. Checks for data leaks or unexpected missing values.
    total_issues = verify_dataset_integrity(df_processed)
    if total_issues =="Failure":
        print(f"⚠️ Detected integrity checks failed! Segregation active.")

        missing_data_row = (df_processed.isin(["Unknown", "unknown"])) | (df_processed.isna())
        non_physical_prop_row = df_processed["Mechanical_Class"]=="Anomalous-Pair"
        merged = missing_data_row.any(axis=1) | non_physical_prop_row
        affected_rows = df_processed[merged]

        df_processed = segregate_data(df_processed, absolute_clean_path, absolute_quarantine_path)
        print(f"✅ Quarantined data located in the Data folder entitled 'Quarantined Anomalies'.")
        total_issues = verify_dataset_integrity(df_processed)
        print("\nAffected rows:")
        print(affected_rows)
    else:
        print(f"✅ Classification verification successful. 100% of rows mapped to known physical categories!")
    
    # 5. Feature Scaling for Neural Network Preparation.
    input_features_to_scale = ["C_Max","Hardness_HB","Yield_strength","Allowable_Contact_Mpa", "Allowable_Bending_Mpa"]
    df_scaled, scaling_manifest = scale_engineering_features(df_processed, input_features_to_scale)
    print(f"✅ The features have been normalized between 0.0 and 1.0 for PINN consumption!")

    # 6. PINN Core-Sizing Training.
    print("Training Physics-Informed Neural Network (PINN) Core Sizing Model...")
    pinn_model = train_pinn_sizing_core(df_scaled, scaling_manifest)
    print("✅ PINN Core training sequence finished. Model asset locked into runtime memory.")

    # 7. Real-World Property Prediction
    engineering_query = {
        "C_Max": 0.20,
        "Hardness_HB": 450.0
    }
    print(f"Evaluating raw engineering boundaries: {engineering_query}")
    query_series = pd.Series(engineering_query)
    physical_predictions = predict_and_unscale_properties(pinn_model, query_series, scaling_manifest)

    # 8. CAD Parameter Bridge Export.
    mechanical_inputs = {
        "torque_nm": 150.0,
        "pinion_teeth": 25,
        "lew_factor_y": 0.35
    }
    # Pass metrics across the parameter bridge to generate 'active_gear_parameters.json'
    export_cad_parameter_payload(physical_predictions, mechanical_inputs)
    print("✅ Parameter pipeline translation complete. 'active_gear_parameters.json' saved to disk.")

    # 9. Autodesk Geometric Automation.
    print("Connecting to Inventor API layer to build geometry canvas...")
    # Executing the 3D generation routine we completed earlier
    cad_success = generate_3d_gear_geometry()
    
    if not cad_success:
        print("❌ Geometry compilation interrupted. Check Autodesk Inventor runtime thread.")

    # 10. Diagnostic Data Visualization.
    for metric in process_metrics:
        saved_path = plot_metallurgical_metrics(
            df=df_processed,
            x_col="Hardness_HB",
            y_metric=metric,
            alloy_class=df_current_class,
        )
        print(f"Generated: {saved_path}")