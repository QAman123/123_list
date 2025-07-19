import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher
import io
import base64


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


def process_excel_file(uploaded_file, progress_bar, status_text):
    """Process a single Excel file"""
    status_text.text(f"Processing file: {uploaded_file.name}")
    progress_bar.progress(0.1)

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

    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file, engine='openpyxl' if uploaded_file.name.endswith('.xlsx') else 'xlrd')
        progress_bar.progress(0.3)

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

        progress_bar.progress(0.5)

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

        progress_bar.progress(0.7)

    except Exception as e:
        st.error(f"Error processing {uploaded_file.name}: {str(e)}")
        return None

    if not all_data:
        st.error("No data extracted from the file")
        return None

    # Create DataFrame from all extracted data
    combined_df = pd.DataFrame(all_data)

    # Clean and convert Final qty to order to numeric
    combined_df['Final qty to order'] = pd.to_numeric(combined_df['Final qty to order'], errors='coerce').fillna(0)

    progress_bar.progress(1.0)
    status_text.text(f"Extracted {len(combined_df)} rows of data from {uploaded_file.name}")

    return combined_df


def combine_similar_items(df, progress_bar, status_text):
    """Combine items with similar names but different sizes"""
    status_text.text("Step 2: Combining similar items...")
    progress_bar.progress(0.1)

    if df is None or df.empty:
        return None, []

    combined_items = []
    combined_sets = []
    processed_indices = set()

    total_items = len(df)
    processed_count = 0

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

        processed_count += 1
        progress_bar.progress(processed_count / total_items)

    # Create new DataFrame with combined items
    combined_df = pd.DataFrame(combined_items)
    status_text.text(f"Combined {len(df)} items into {len(combined_df)} items ({len(combined_sets)} sets combined)")

    return combined_df, combined_sets


def add_price_break_analysis(df, progress_bar, status_text):
    """Add price break tier analysis to the dataframe"""
    status_text.text("Step 3: Calculating price break tiers...")
    progress_bar.progress(0.1)

    if df is None or df.empty:
        return None

    # Create a copy to avoid modifying the original
    result_df = df.copy()

    # Calculate price break tier for each item (returns tuple now)
    price_break_results = result_df.apply(calculate_price_break, axis=1)
    progress_bar.progress(0.5)

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

    progress_bar.progress(1.0)
    status_text.text("Price break analysis completed")
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


def add_cost_analysis(df, progress_bar, status_text):
    """Add cost analysis columns to the dataframe"""
    status_text.text("Step 4: Calculating cost analysis...")
    progress_bar.progress(0.1)

    if df is None or df.empty:
        return None

    result_df = df.copy()

    # Initialize new columns
    result_df['cost_current_PB'] = None
    result_df['cost_higher_PB'] = None
    result_df['cost_difference'] = None
    result_df['higher_tier_qty_required'] = None

    total_items = len(result_df)

    for idx, (i, row) in enumerate(result_df.iterrows()):
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

        progress_bar.progress((idx + 1) / total_items)

    status_text.text("Cost analysis completed")
    return result_df


def get_download_link(df, filename):
    """Generate a download link for the dataframe"""
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download {filename}</a>'
    return href


def generate_summary_report(results):
    """Generate a text summary report of the analysis"""
    final_df = results['final_df']
    combined_sets = results['combined_sets']
    original_df = results['original_df']
    original_filename = results.get('original_filename', 'analysis')

    report = []
    report.append("=" * 60)
    report.append("PRICE BREAK ANALYSIS SUMMARY REPORT")
    report.append("=" * 60)
    report.append(f"Original filename: {original_filename}")
    report.append(f"Analysis date: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append("")

    # Summary
    report.append("SUMMARY:")
    report.append(f"- Original nr items: {len(original_df)}")
    report.append(f"- Nr items after combining sizes: {len(final_df)}")
    report.append(f"- {len(original_df) - len(final_df)} items combined into {len(combined_sets)} items")
    report.append("")

    # Combined sets
    # if combined_sets:
    #     report.append("COMBINED SETS:")
    #     for i, combo in enumerate(combined_sets, 1):
    #         report.append(f"\n{i}. Combined into: '{combo['combined_name']}'")
    #         report.append(f"   Total quantity: {combo['total_qty']}")
    #         report.append("   Original items:")
    #         for orig_name in combo['original_names']:
    #             report.append(f"   - {orig_name}")
    # else:
    #     report.append("No items were combined (no similar items found)")
    #
    # report.append("")

    # Price break analysis
    report.append("PRICE BREAK ANALYSIS:")
    price_break_counts = final_df['Price Break Tier'].value_counts()
    for tier, count in price_break_counts.items():
        report.append(f"- {tier}: {count} items")

    report.append("")

    # PB percentage analysis
    pb_items = final_df[final_df['Highest PB Percentage'].notna()]
    if not pb_items.empty:
        pb_items = pb_items.copy()
        pb_items['Highest PB Percentage'] = pd.to_numeric(pb_items['Highest PB Percentage'], errors='coerce')
        pb_items = pb_items[pb_items['Highest PB Percentage'].notna()]

        if not pb_items.empty:
            report.append("ITEMS AT HIGHEST PRICE BREAK (PERCENTAGE ANALYSIS):")
            report.append(f"- Total items at highest break: {len(pb_items)}")
            report.append(f"- Average percentage of highest PB: {pb_items['Highest PB Percentage'].mean():.2f}%")

            report.append(f"\nALL items in their highest price break:")
            all_pb = pb_items.sort_values('Highest PB Percentage', ascending=False)[
                ['Article name', 'Final qty to order', 'Price Break Tier', 'Highest PB Percentage']]
            for _, row in all_pb.iterrows():
                report.append(
                    f"- {row['Article name']}: {row['Final qty to order']:.0f} qty ({row['Highest PB Percentage']:.2f}% of PB{row['Price Break Tier']})")

    report.append("")

    # Cost analysis
    cost_analysis_items = final_df[
        (final_df['cost_current_PB'].notna()) &
        (final_df['cost_higher_PB'].notna()) &
        (final_df['Price Break Tier'] != 'BELOW_MOQ') &
        (final_df['Price Break Tier'] != 'NO_PRICE_BREAKS')
        ]

    if not cost_analysis_items.empty:
        report.append("COST ANALYSIS FOR ITEMS NOT AT HIGHEST PRICE BREAK:")
        report.append(f"- Total items analyzed: {len(cost_analysis_items)}")

        lower_cost_items = cost_analysis_items[
            cost_analysis_items['cost_higher_PB'] < cost_analysis_items['cost_current_PB']]

        if not lower_cost_items.empty:
            report.append(
                f"\nItems where higher price break costs LESS than current price break - POTENTIAL SAVINGS ({len(lower_cost_items)} items):")
            total_savings = 0

            for _, row in lower_cost_items.iterrows():
                article_name = row.get('Article name', '')
                current_qty = int(row.get('Final qty to order', 0))
                pb_tier = float(row.get('Price Break Tier', 0))
                cost_current = float(row.get('cost_current_PB', 0))
                higher_qty = float(row.get('higher_tier_qty_required', 0))
                #higher_qty = higher_qty + 0.01
                cost_higher = float(row.get('cost_higher_PB', 0))
                cost_diff = float(row.get('cost_difference', 0))

                savings = abs(row['cost_difference'])
                total_savings += savings
                report.append(f"- {row['Article name']}:")
                report.append(
                    f"  Current ordered: {row['Final qty to order']:.0f} qty at PB{row['Price Break Tier']} = €{row['cost_current_PB']:.0f}")
                report.append(
                    f"  Higher: {row['higher_tier_qty_required']:.0f} qty at next PB = €{row['cost_higher_PB']:.0f}")
                report.append(
                    f"  Opco will save: €{(cost_current - ((cost_higher / higher_qty) * current_qty)):.0f}")
                report.append(
                    f"  Global needs to spend: -€{((cost_higher / higher_qty) * (higher_qty - current_qty)):.0f}")
                report.append(
                    f"  Global takes on stock: {(higher_qty - current_qty):.0f} qty")
                report.append(f"  SAVINGS: €{savings:.2f}")
                report.append("")


            report.append(f"TOTAL POTENTIAL SAVINGS: €{total_savings:.0f}")
        else:
            report.append("No items found where higher price break would result in savings.")

    report.append("")
    report.append("=" * 60)
    report.append("END OF REPORT - FCH")
    report.append("=" * 60)

    return "\n".join(report)


def get_base_filename(uploaded_files):
    """Get base filename from uploaded files"""
    if len(uploaded_files) == 1:
        # Single file - use its name without extension
        return uploaded_files[0].name.rsplit('.', 1)[0]
    else:
        # Multiple files - use a generic name with count
        return f"combined_{len(uploaded_files)}_files"


def main():
    st.set_page_config(page_title="Price Break Analysis Tool", layout="wide")

    st.title("🔍 123-List Analysis Tool")
    st.markdown("Upload your Excel files to analyze price breaks and find cost optimization opportunities")

    # File upload
    uploaded_files = st.file_uploader(
        "Choose Excel files",
        accept_multiple_files=True,
        type=['xlsx', 'xls'],
        help="Upload one or more Excel files containing pricing data"
    )

    if uploaded_files:
        # Initialize session state for results
        if 'results' not in st.session_state:
            st.session_state.results = None

        # Process files button
        if st.button("🚀 Process Files", type="primary"):
            with st.spinner("Processing files..."):
                # Create progress indicators
                progress_container = st.container()
                with progress_container:
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                all_data = []

                # Process each file
                for uploaded_file in uploaded_files:
                    df = process_excel_file(uploaded_file, progress_bar, status_text)
                    if df is not None:
                        all_data.append(df)

                if not all_data:
                    st.error("No data was extracted from any files")
                    return

                # Combine all data
                combined_original_df = pd.concat(all_data, ignore_index=True)

                # Step 2: Combine similar items
                combined_df, combined_sets = combine_similar_items(combined_original_df, progress_bar, status_text)

                if combined_df is None:
                    st.error("Failed to combine similar items")
                    return

                # Step 3: Add price break analysis
                final_df = add_price_break_analysis(combined_df, progress_bar, status_text)

                # Step 4: Add cost analysis
                final_df = add_cost_analysis(final_df, progress_bar, status_text)

                # Store results in session state
                base_filename = get_base_filename(uploaded_files)
                st.session_state.results = {
                    'original_df': combined_original_df,
                    'final_df': final_df,
                    'combined_sets': combined_sets,
                    'original_filename': base_filename
                }

                progress_container.empty()
                st.success("✅ Processing completed!")

        # Display results if available
        if st.session_state.results:
            results = st.session_state.results
            final_df = results['final_df']
            combined_sets = results['combined_sets']
            original_df = results['original_df']

            # Summary metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Original nr of items", len(original_df))
            with col2:
                st.metric("Nr of items after combining sizes", len(final_df))
            # with col3:
            #     st.metric("Items Saved", len(original_df) - len(final_df))
            # with col4:
            #     st.metric("Sets Combined", len(combined_sets))

            # Combined sets (in expandable section)
            if combined_sets:
                with st.expander(f"🔄 Combined Sets ({len(combined_sets)} total)", expanded=False):
                    for i, combo in enumerate(combined_sets, 1):
                        st.markdown(f"**{i}. Combined into:** {combo['combined_name']}")
                        st.markdown(f"**Total quantity:** {combo['total_qty']}")
                        st.markdown("**Original items:**")
                        for orig_name in combo['original_names']:
                            st.markdown(f"   - {orig_name}")
                        st.markdown("---")
            st.write("")
            # Download buttons
            st.markdown("### 📥 Download Results")
            base_filename = results.get('original_filename', 'analysis')

            col1, col2 = st.columns(2)

            with col1:
                csv = final_df.to_csv(index=False)
                st.download_button(
                    label="📊 Download Full Data Analysis",
                    data=csv,
                    file_name=f"{base_filename}_full_data_analysis.csv",
                    mime="text/csv",
                    help="Complete dataset with all analysis columns"
                )

            with col2:
                summary_report = generate_summary_report(results)
                st.download_button(
                    label="📋 Download Summary Report",
                    data=summary_report,
                    file_name=f"{base_filename}_summary_report.txt",
                    mime="text/plain",
                    help="Text summary of all analysis results"
                )

            st.write("")

            # Price Break Analysis
            st.markdown("### 📊 Price Break Analysis")
            price_break_counts = final_df['Price Break Tier'].value_counts()

            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Price Break Distribution:**")
                for tier, count in price_break_counts.items():
                    st.markdown(f"- Pricebreak {tier}: {count} items")

            with col2:
                # Chart of price break distribution
                st.bar_chart(price_break_counts)

            # Items at Highest Price Break
            pb_items = final_df[final_df['Highest PB Percentage'].notna()]
            if not pb_items.empty:
                st.markdown("### 📈 Items at Highest Price Break (Percentage Analysis)")

                # Convert to numeric and handle any conversion issues
                pb_items = pb_items.copy()
                pb_items['Highest PB Percentage'] = pd.to_numeric(pb_items['Highest PB Percentage'], errors='coerce')
                pb_items = pb_items[pb_items['Highest PB Percentage'].notna()]

                if not pb_items.empty:
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total items at highest break", len(pb_items))
                    with col2:
                        st.metric("Average % of highest PB", f"{pb_items['Highest PB Percentage'].mean():.2f}%")
                    #with col3:
                        # above_100 = len(pb_items[pb_items['Highest PB Percentage'] > 100])
                        # st.metric("Items above 100% of highest PB", above_100)

                    # Show all items with their percentages
                    st.markdown("#### 🎯 ALL items with highest percentages of their price break:")
                    all_pb = pb_items.sort_values('Highest PB Percentage', ascending=False)[
                        ['Article name', 'Final qty to order', 'Price Break Tier', 'Highest PB Percentage']]

                    # Display as a table
                    st.dataframe(
                        all_pb,
                        column_config={
                            "Article name": st.column_config.TextColumn("Article Name", width="large"),
                            "Final qty to order": st.column_config.NumberColumn("Quantity", format="%d"),
                            "Price Break Tier": st.column_config.NumberColumn("Pricebreak"),
                            "Highest PB Percentage": st.column_config.NumberColumn("Pricebreak Percentage", format="%.2f%%")
                        },
                        hide_index=True
                    )

            # Cost Analysis
            cost_analysis_items = final_df[
                (final_df['cost_current_PB'].notna()) &
                (final_df['cost_higher_PB'].notna()) &
                (final_df['Price Break Tier'] != 'BELOW_MOQ') &
                (final_df['Price Break Tier'] != 'NO_PRICE_BREAKS')
                ]

            if not cost_analysis_items.empty:
                st.markdown("### 💰 Cost Analysis for Items Not at Highest Price Break")
                st.metric("Total items analyzed", len(cost_analysis_items))

                # Items where higher price break costs less
                lower_cost_items = cost_analysis_items[
                    cost_analysis_items['cost_higher_PB'] < cost_analysis_items['cost_current_PB']]

                if not lower_cost_items.empty:
                    st.markdown(
                        f"#### 💡 Items where higher price break costs LESS - POTENTIAL SAVINGS ({len(lower_cost_items)} items)")

                    # Create a detailed view for savings
                    savings_data = []
                    for _, row in lower_cost_items.iterrows():

                        article_name = row.get('Article name', '')
                        current_qty = int(row.get('Final qty to order', 0))
                        pb_tier = float(row.get('Price Break Tier', 0))
                        cost_current = float(row.get('cost_current_PB', 0))
                        higher_qty = float(row.get('higher_tier_qty_required', 0))
                        cost_higher = float(row.get('cost_higher_PB', 0))
                        cost_diff = float(row.get('cost_difference', 0))

                        savings_data.append({
                            'Article Name': row['Article name'],
                            'Current Qty': int(row['Final qty to order']),
                            'Current PB': f"{row['Price Break Tier']}",
                            'Current Cost': f"€{row['cost_current_PB']:.0f}",
                            'Higher Qty Required': int(row['higher_tier_qty_required']),
                            'Higher PB Cost': f"€{row['cost_higher_PB']:.0f}",
                            'Savings for Opco': f"€{(cost_current-((cost_higher/higher_qty)*current_qty)):.0f}",
                            'Spend for Global': f"-€{((cost_higher/higher_qty)*(higher_qty-current_qty)):.0f}",
                            'Stock Qty Global': f"{(higher_qty-current_qty):.0f}",
                            'SAVINGS': f"€{abs(row['cost_difference']):.0f}"
                        })

                    savings_df = pd.DataFrame(savings_data)
                    st.dataframe(
                        savings_df,
                        column_config={
                            "Article Name": st.column_config.TextColumn("Article Name", width="large"),
                            "Current Qty": st.column_config.TextColumn("Current Qty"),
                            "Current PB": st.column_config.TextColumn("Current Pricebreak"),
                            "Current Cost": st.column_config.TextColumn("Cost at current Pricebreak"),
                            "Higher Qty Required": st.column_config.TextColumn("Qty at higher Pricebreak"),
                            "Higher PB Cost": st.column_config.TextColumn("Cost at higher Pricebreak"),
                            "Savings for Opco": st.column_config.TextColumn("Savings for Opco"),
                            "Spend for Global": st.column_config.TextColumn("Spend for Global"),
                            "Stock Qty Global": st.column_config.TextColumn("Stock Qty Global"),
                            "SAVINGS": st.column_config.TextColumn("💰 SAVINGS OVERALL", width="medium")
                        },
                        hide_index=True
                    )
                    st.write("Clarification: Global needs to take and pay for the extra quantity to achieve higher pricebreak. Savings for the Opcos will be larger than Global spend.")
                    # Total potential savings
                    total_savings = abs(lower_cost_items['cost_difference'].sum())
                    st.success(f"🎉 **Total Potential Savings: €{total_savings:.0f}**")

    else:
        st.info("👆 Please upload Excel files to begin analysis")

        # Show example of expected file structure
        with st.expander("📋 Expected File Structure", expanded=False):
            st.markdown("""
            Your Excel files should contain the following columns:
            - **Brand**: Product brand
            - **Article no.**: Article number/SKU
            - **Article name**: Product name (sizes will be automatically detected)
            - **Final qty to order**: Quantity to order
            - **Price break 1-5**: Quantity thresholds for each price break
            - **Price break 1-5 purchase price**: Corresponding prices for each break

            The tool will automatically:
            - Combine items with the same name but different sizes (XS, S, M, L, XL, etc.)
            - Calculate which price break tier each item falls into
            - Identify opportunities for cost savings by ordering higher quantities
            """)


if __name__ == "__main__":
    main()