import pandas as pd
import os
import glob
import re
from difflib import SequenceMatcher


def similarity(a, b):
    """Calculate similarity between two strings"""
    return SequenceMatcher(None, a, b).ratio()


def extract_size_from_name(name):
    """Extract ONLY clothing size information from article name"""
    # Only these specific clothing sizes can be combined
    allowed_sizes = ['XXS', 'XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL', '3XL']

    # Look for these sizes at the end of the name (with optional punctuation/spaces)
    size_pattern = r'\b(' + '|'.join(allowed_sizes) + r')\b\s*[,\-\s]*$'

    sizes_found = []
    name_without_size = name

    # Find all matches for allowed sizes at the end
    matches = list(re.finditer(size_pattern, name, re.IGNORECASE))

    if matches:
        # Take the last match (rightmost)
        match = matches[-1]
        size = match.group(1).upper()
        if size in allowed_sizes:
            sizes_found.append(size)
            # Remove the size from the end of the name
            name_without_size = name[:match.start()].strip()
            # Clean up trailing punctuation
            name_without_size = re.sub(r'[,\-\s]+$', '', name_without_size).strip()

    return name_without_size, sizes_found


def can_combine_items(name1, name2):
    """Check if two items can be combined - only if exact same name except for clothing size at the end"""
    base_name1, sizes1 = extract_size_from_name(name1)
    base_name2, sizes2 = extract_size_from_name(name2)

    # Items can only be combined if:
    # 1. They have exactly the same base name (case insensitive)
    # 2. Both have exactly one clothing size
    # 3. The sizes are different

    if not sizes1 or not sizes2:
        return False, base_name1  # No sizes found, can't combine

    if len(sizes1) != 1 or len(sizes2) != 1:
        return False, base_name1  # Multiple sizes or no sizes, can't combine

    # Check if base names are exactly the same (ignoring case and extra spaces)
    base1_clean = re.sub(r'\s+', ' ', base_name1.strip().lower())
    base2_clean = re.sub(r'\s+', ' ', base_name2.strip().lower())

    if base1_clean == base2_clean and sizes1[0] != sizes2[0]:
        return True, base_name1 if len(base_name1) >= len(base_name2) else base_name2

    return False, base_name1


def calculate_price_break(row):
    """Calculate which price break tier applies based on Final qty to order"""
    qty = row['Final qty to order']

    # Get all price break quantities (convert to numeric, treat 'na' and NaN as None)
    price_breaks = []
    for i in range(1, 6):  # Price breaks 1-5
        col_name = f'Price break {i}'
        if col_name in row:
            value = row[col_name]
            if pd.notna(value) and str(value).strip().lower() != 'na' and str(value).strip() != '':
                try:
                    price_breaks.append((i, float(value)))
                except (ValueError, TypeError):
                    continue

    # Sort price breaks by quantity (ascending)
    price_breaks.sort(key=lambda x: x[1])

    # If no valid price breaks found
    if not price_breaks:
        return "NO_PRICE_BREAKS", None

    # Check if quantity is below the first price break (MOQ)
    if qty < price_breaks[0][1]:
        return "BELOW_MOQ", None

    # Find the appropriate price break
    selected_tier = None
    is_highest_break = False

    for i in range(len(price_breaks)):
        current_break_num, current_qty = price_breaks[i]

        # If this is the last price break, or qty is less than the next break
        if i == len(price_breaks) - 1:
            selected_tier = str(current_break_num)
            is_highest_break = True
            break
        else:
            next_break_num, next_qty = price_breaks[i + 1]
            if qty < next_qty:
                selected_tier = str(current_break_num)
                # Check if this is the highest available break (next breaks are 'na')
                is_highest_break = (current_break_num == 5 or
                                    not any(pb[0] > current_break_num for pb in price_breaks))
                break

    # If we get here without setting selected_tier, use the highest price break
    if selected_tier is None:
        selected_tier = str(price_breaks[-1][0])
        is_highest_break = True

    return selected_tier, is_highest_break


def calculate_highest_pb_percentage(row, highest_tier):
    """Calculate what percentage the final qty is of the highest price break tier"""
    qty = row['Final qty to order']

    # Get the value for the highest price break tier
    pb_col = f'Price break {highest_tier}'
    if pb_col in row:
        pb_value = row[pb_col]
        if pd.notna(pb_value) and str(pb_value).strip().lower() != 'na' and str(pb_value).strip() != '':
            try:
                pb_qty = float(pb_value)
                if pb_qty > 0:
                    percentage = (qty / pb_qty) * 100
                    return round(percentage, 2)
            except (ValueError, TypeError):
                pass

    return None


def process_excel_files(folder_path):
    """Process all Excel files in the specified folder"""

    # Get all .xls and .xlsx files in the folder
    excel_files = []
    excel_files.extend(glob.glob(os.path.join(folder_path, "*.xls")))
    excel_files.extend(glob.glob(os.path.join(folder_path, "*.xlsx")))

    if not excel_files:
        print(f"No Excel files found in folder: {folder_path}")
        return None

    print(f"Found {len(excel_files)} Excel files to process")

    # Define the columns we want to extract
    required_columns = [
        'Brand', 'Article no.', 'Article name', 'Final qty to order',
        'Price break 1', 'Price break 1\n\npurchase price\nInco term 1',
        'Price break 2', 'Price break 2\n\npurchase price\nInco term 1',
        'Price break 3', 'Price break 3\n\npurchase price\nInco term 1',
        'Price break 4', 'Price break 4\n\npurchase price\nInco term 1',
        'Price break 5', 'Price break 5\n\npurchase price\nInco term 1'
    ]

    all_data = []

    # Process each Excel file
    for file_path in excel_files:
        print(f"Processing: {os.path.basename(file_path)}")

        try:
            # Try to read the Excel file
            df = pd.read_excel(file_path, engine='openpyxl' if file_path.endswith('.xlsx') else 'xlrd')

            # Create a dictionary to store the extracted data
            extracted_data = {}

            # Initialize all columns with empty lists
            for col in required_columns:
                extracted_data[col] = []

            # Find matching columns (case-insensitive and flexible matching)
            column_mapping = {}
            df_columns_lower = [col.lower().strip().replace('\n', ' ').replace('  ', ' ') for col in df.columns]
            used_columns = set()  # Track which columns have been matched

            for req_col in required_columns:
                req_col_lower = req_col.lower().strip().replace('\n', ' ').replace('  ', ' ')

                # Try exact match first
                exact_match_found = False
                for i, df_col_lower in enumerate(df_columns_lower):
                    if req_col_lower == df_col_lower and df.columns[i] not in used_columns:
                        column_mapping[req_col] = df.columns[i]
                        used_columns.add(df.columns[i])
                        exact_match_found = True
                        break

                if not exact_match_found:
                    # Try similarity matching but ensure we don't reuse columns
                    best_match = None
                    best_similarity = 0

                    for i, df_col in enumerate(df.columns):
                        if df_col in used_columns:
                            continue

                        df_col_normalized = df_col.lower().strip().replace('\n', ' ').replace('  ', ' ')
                        sim = similarity(req_col_lower, df_col_normalized)

                        if sim > 0.8 and sim > best_similarity:
                            best_similarity = sim
                            best_match = df_col

                    if best_match:
                        column_mapping[req_col] = best_match
                        used_columns.add(best_match)

            # Extract data for each row
            for idx, row in df.iterrows():
                row_data = {}
                has_required_data = False

                for req_col in required_columns:
                    if req_col in column_mapping:
                        value = row[column_mapping[req_col]]
                        if pd.notna(value) and str(value).strip() != '':
                            row_data[req_col] = value
                            if req_col in ['Brand', 'Article no.', 'Article name']:
                                has_required_data = True
                        else:
                            row_data[req_col] = 'na'
                    else:
                        row_data[req_col] = 'na'

                # Only add rows that have at least some required data
                if has_required_data:
                    all_data.append(row_data)

        except Exception as e:
            print(f"Error processing {file_path}: {str(e)}")
            continue

    if not all_data:
        print("No data extracted from any files")
        return None

    # Create DataFrame from all extracted data
    combined_df = pd.DataFrame(all_data)

    # Clean and convert Final qty to order to numeric
    combined_df['Final qty to order'] = pd.to_numeric(combined_df['Final qty to order'], errors='coerce').fillna(0)

    print(f"Extracted {len(combined_df)} rows of data")
    return combined_df


def combine_similar_items(df):
    """Combine items with similar names but different sizes"""

    if df is None or df.empty:
        return None, []

    combined_items = []
    combined_sets = []
    processed_indices = set()

    for i in range(len(df)):
        if i in processed_indices:
            continue

        current_item = df.iloc[i].copy()
        current_name = str(current_item['Article name'])
        group_indices = [i]
        group_sizes = []

        # Extract size from current item
        base_name, sizes = extract_size_from_name(current_name)
        if sizes:
            group_sizes.extend(sizes)

        # Look for similar items
        for j in range(i + 1, len(df)):
            if j in processed_indices:
                continue

            other_item = df.iloc[j]
            other_name = str(other_item['Article name'])

            # Check if items can be combined
            can_combine, common_base = can_combine_items(current_name, other_name)

            if can_combine:
                # Extract sizes from the other item
                _, other_sizes = extract_size_from_name(other_name)
                if other_sizes:
                    group_sizes.extend(other_sizes)

                group_indices.append(j)
                processed_indices.add(j)

                # Add quantities
                current_item['Final qty to order'] += other_item['Final qty to order']

        processed_indices.add(i)

        # Update the article name if items were combined
        if len(group_indices) > 1:
            # Remove duplicates and sort sizes
            unique_sizes = list(set(group_sizes))

            # Sort sizes in the proper order
            size_order = ['XXS', 'XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL', '3XL']
            sorted_sizes = []

            for size in size_order:
                if size in unique_sizes:
                    sorted_sizes.append(size)

            if sorted_sizes:
                current_item['Article name'] = f"{base_name} ({', '.join(sorted_sizes)})"
            else:
                current_item['Article name'] = base_name

            # Record what was combined
            original_names = [str(df.iloc[idx]['Article name']) for idx in group_indices]
            combined_sets.append({
                'combined_name': current_item['Article name'],
                'original_names': original_names,
                'total_qty': current_item['Final qty to order']
            })

        combined_items.append(current_item)

    # Create new DataFrame with combined items
    combined_df = pd.DataFrame(combined_items)

    return combined_df, combined_sets


def add_price_break_analysis(df):
    """Add price break tier analysis to the dataframe"""
    if df is None or df.empty:
        return None

    # Create a copy to avoid modifying the original
    result_df = df.copy()

    # Calculate price break tier for each item (returns tuple now)
    price_break_results = result_df.apply(calculate_price_break, axis=1)

    # Separate the tier and highest_break flag
    result_df['Price Break Tier'] = [result[0] for result in price_break_results]
    is_highest_break = [result[1] for result in price_break_results]

    # Calculate percentage for items at highest break (excluding BELOW_MOQ)
    result_df['Highest PB Percentage'] = None

    # Use enumerate to get sequential index for is_highest_break list
    for idx, (i, row) in enumerate(result_df.iterrows()):
        tier = row['Price Break Tier']
        if tier not in ['BELOW_MOQ', 'NO_PRICE_BREAKS'] and is_highest_break[idx]:
            # Calculate percentage of the actual tier they're at (which is the highest available)
            highest_tier = int(tier)
            pb_percentage = calculate_highest_pb_percentage(row, highest_tier)
            result_df.at[i, 'Highest PB Percentage'] = pb_percentage

    return result_df


def get_price_for_tier(row, tier):
    """Get the price for a specific price break tier"""
    if tier in [1, 2, 3, 4, 5]:
        price_col = f'Price break {tier}\n\npurchase price\nInco term 1'
        if price_col in row:
            value = row[price_col]
            if pd.notna(value) and str(value).strip().lower() != 'na' and str(value).strip() != '':
                try:
                    return float(value)
                except (ValueError, TypeError):
                    pass
    return None


def get_next_higher_tier(row, current_tier):
    """Find the next higher available price break tier"""
    if current_tier in ['BELOW_MOQ', 'NO_PRICE_BREAKS']:
        return None

    current_tier_num = int(current_tier)

    # Check tiers above the current one
    for tier in range(current_tier_num + 1, 6):
        price_col = f'Price break {tier}'
        if price_col in row:
            value = row[price_col]
            if pd.notna(value) and str(value).strip().lower() != 'na' and str(value).strip() != '':
                try:
                    tier_qty = float(value)
                    if tier_qty > 0:
                        return tier
                except (ValueError, TypeError):
                    continue
    return None


def add_cost_analysis(df):
    """Add cost analysis columns to the dataframe"""
    if df is None or df.empty:
        return None

    result_df = df.copy()

    # Initialize new columns
    result_df['cost_current_PB'] = None
    result_df['cost_higher_PB'] = None
    result_df['cost_difference'] = None
    result_df['higher_tier_qty_required'] = None

    for i, row in result_df.iterrows():
        tier = row['Price Break Tier']
        qty = row['Final qty to order']

        # Skip items that are BELOW_MOQ or have NO_PRICE_BREAKS
        if tier in ['BELOW_MOQ', 'NO_PRICE_BREAKS']:
            continue

        # Get current tier price
        current_tier_num = int(tier)
        current_price = get_price_for_tier(row, current_tier_num)

        if current_price is not None:
            current_cost = qty * current_price
            result_df.at[i, 'cost_current_PB'] = current_cost

            # Find next higher tier
            next_tier = get_next_higher_tier(row, tier)

            if next_tier is not None:
                # Get the quantity required for the higher tier
                higher_tier_qty_col = f'Price break {next_tier}'
                if higher_tier_qty_col in row:
                    higher_tier_qty = row[higher_tier_qty_col]
                    if pd.notna(higher_tier_qty) and str(higher_tier_qty).strip().lower() != 'na':
                        try:
                            higher_tier_qty = float(higher_tier_qty)
                            result_df.at[i, 'higher_tier_qty_required'] = higher_tier_qty

                            # Get price for higher tier
                            higher_price = get_price_for_tier(row, next_tier)

                            if higher_price is not None:
                                # Calculate cost if we ordered the higher tier quantity
                                higher_cost = higher_tier_qty * higher_price
                                result_df.at[i, 'cost_higher_PB'] = higher_cost

                                # Calculate difference
                                cost_diff = higher_cost - current_cost
                                result_df.at[i, 'cost_difference'] = cost_diff
                        except (ValueError, TypeError):
                            pass

    return result_df


def main():
    folder_path = "123_files"

    # Check if folder exists
    if not os.path.exists(folder_path):
        print(f"Folder '{folder_path}' does not exist!")
        return

    # Process Excel files
    print("Step 1: Processing Excel files...")
    df = process_excel_files(folder_path)

    if df is None:
        return

    # Export original data
    df.to_csv("original_data.csv", index=False)
    print(f"Original data exported to 'original_data.csv' ({len(df)} items)")

    # Combine similar items
    print("\nStep 2: Combining similar items...")
    combined_df, combined_sets = combine_similar_items(df)

    if combined_df is None:
        return

    # Add price break analysis
    print("\nStep 3: Calculating price break tiers...")
    final_df = add_price_break_analysis(combined_df)

    # Add cost analysis
    print("\nStep 4: Calculating cost analysis...")
    final_df = add_cost_analysis(final_df)

    # Print combined sets
    if combined_sets:
        print("\nCombined sets:")
        for i, combo in enumerate(combined_sets, 1):
            print(f"\n{i}. Combined into: '{combo['combined_name']}'")
            print(f"   Total quantity: {combo['total_qty']}")
            print("   Original items:")
            for orig_name in combo['original_names']:
                print(f"   - {orig_name}")
    else:
        print("No items were combined (no similar items found)")

    # Print price break analysis summary
    print(f"\nPrice Break Analysis:")
    price_break_counts = final_df['Price Break Tier'].value_counts()
    for tier, count in price_break_counts.items():
        print(f"- {tier}: {count} items")

    # Print PB percentage analysis (ALL items, not just top 5)
    pb_items = final_df[final_df['Highest PB Percentage'].notna()]
    if not pb_items.empty:
        print(f"\nItems at Highest Price Break (Percentage Analysis):")
        print(f"- Total items at highest break: {len(pb_items)}")

        # Convert to numeric and handle any conversion issues
        pb_items = pb_items.copy()
        pb_items['Highest PB Percentage'] = pd.to_numeric(pb_items['Highest PB Percentage'], errors='coerce')

        # Remove any rows where conversion failed
        pb_items = pb_items[pb_items['Highest PB Percentage'].notna()]

        if not pb_items.empty:
            print(f"- Average percentage of highest PB: {pb_items['Highest PB Percentage'].mean():.2f}%")
            print(f"- Items above 100% of highest PB: {len(pb_items[pb_items['Highest PB Percentage'] > 100])}")

            # Show ALL items with their percentages
            print(f"\nALL items with highest percentages of their price break:")
            all_pb = pb_items.sort_values('Highest PB Percentage', ascending=False)[
                ['Article name', 'Final qty to order', 'Price Break Tier', 'Highest PB Percentage']]
            for _, row in all_pb.iterrows():
                print(
                    f"- {row['Article name']}: {row['Final qty to order']} qty ({row['Highest PB Percentage']:.2f}% of PB{row['Price Break Tier']})")
        else:
            print("- No valid percentage data found after conversion")

    # Print cost analysis for items not at highest price break
    cost_analysis_items = final_df[
        (final_df['cost_current_PB'].notna()) &
        (final_df['cost_higher_PB'].notna()) &
        (final_df['Price Break Tier'] != 'BELOW_MOQ') &
        (final_df['Price Break Tier'] != 'NO_PRICE_BREAKS')
        ]

    if not cost_analysis_items.empty:
        print(f"\nCost Analysis for Items Not at Highest Price Break:")
        print(f"- Total items analyzed: {len(cost_analysis_items)}")



        lower_cost_items = cost_analysis_items[
            cost_analysis_items['cost_higher_PB'] < cost_analysis_items['cost_current_PB']]

        if not lower_cost_items.empty:
            print(
                f"\nItems where higher price break costs LESS than current price break - POTENTIAL SAVINGS ({len(lower_cost_items)} items):")
            for _, row in lower_cost_items.iterrows():
                print(f"- {row['Article name']}:")
                print(
                    f"  Current: {row['Final qty to order']} qty at PB{row['Price Break Tier']} = ${row['cost_current_PB']:.2f}")
                print(f"  Higher: {row['higher_tier_qty_required']} qty at next PB = ${row['cost_higher_PB']:.2f}")
                print(f"  SAVINGS: ${abs(row['cost_difference']):.2f}")
                print()

    # Export final data with cost analysis
    final_df.to_csv("final_data_with_price_breaks.csv", index=False)
    print(
        f"\nFinal data with price break and cost analysis exported to 'final_data_with_price_breaks.csv' ({len(final_df)} items)")

    # Show summary
    print(f"\nSummary:")
    print(f"- Original items: {len(df)}")
    print(f"- Combined items: {len(combined_df)}")
    print(f"- Items saved: {len(df) - len(combined_df)}")
    print(f"- Sets combined: {len(combined_sets)}")

if __name__ == "__main__":
    main()