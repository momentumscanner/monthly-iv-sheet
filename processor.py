
import pandas as pd
import zipfile
import os
import io

class Processor:

    def __init__(self, zip_path, options_screener_path=None):
        self.zip_path = zip_path
        self.options_screener_path = options_screener_path
        self.log_callback = print  # Default to print

    def set_log_callback(self, callback):
        self.log_callback = callback

    def log(self, message):
        if self.log_callback:
            self.log_callback(message)

    def get_daywise_data(self):
        """Reads the Daywise Price and OI Summary Excel file and returns a dictionary of Symbol -> {IV, IVP, IVR}."""
        if not self.options_screener_path or not os.path.exists(self.options_screener_path):
            return {}
        
        try:
            self.log(f"Reading Daywise Report: {os.path.basename(self.options_screener_path)}")
            # Read Excel file
            df = pd.read_excel(self.options_screener_path)
            
            # Ensure required columns exist
            # Expected columns: Symbol, IV, IVP, IVR
            required_cols = ['Symbol', 'IV', 'IVP', 'IVR']
            missing = [col for col in required_cols if col not in df.columns]
            
            if missing:
                self.log(f"Warning: Daywise Report missing columns: {missing}")
                return {}
            
            # Create dictionary: Symbol -> {IV, IVP, IVR}
            data_map = {}
            for _, row in df.iterrows():
                symbol = row['Symbol']
                data_map[symbol] = {
                    'IV': row['IV'],
                    'IVP': row['IVP'],
                    'IVR': row['IVR']
                }
            
            self.log(f"Loaded data for {len(data_map)} symbols from Daywise Report")
            return data_map
            
        except Exception as e:
            self.log(f"Error reading Daywise Report: {e}")
            return {}

    def get_expiries(self):
        """Quickly scan the file to get unique Expiry dates for FUTSTK."""
        try:
            with zipfile.ZipFile(self.zip_path, 'r') as z:
                csv_filename = [n for n in z.namelist() if n.lower().endswith('.csv') and 'bhav' in n.lower()][0]
                with z.open(csv_filename) as f:
                    # Read unique expiries - efficient approach?
                    # Just read the whole thing, pandas is fast enough for 5MB
                    df = pd.read_csv(f)
            
            # Normalize columns locally for this check
            col_map = {
                'FinInstrmTp': 'INSTRUMENT',
                'TckrSymb': 'SYMBOL',
                'XpryDt': 'EXPIRY_DT',
                'StrkPric': 'STRIKE_PR',
                'OptnTp': 'OPTION_TYP',
                'ClsPric': 'CLOSE'
            }
            df.rename(columns=col_map, inplace=True)
            
            if 'INSTRUMENT' in df.columns:
                 df['INSTRUMENT'] = df['INSTRUMENT'].replace({'STF': 'FUTSTK', 'STO': 'OPTSTK', 'IDF': 'FUTIDX', 'IDO': 'OPTIDX'})
            
            if 'EXPIRY_DT' not in df.columns or 'INSTRUMENT' not in df.columns:
                return []

            # Get futures expiries
            expiries = df[df['INSTRUMENT'] == 'FUTSTK']['EXPIRY_DT'].unique()
            
            # Sort by date
            # Try to parse dates for sorting
            try:
                # Handle YYYY-MM-DD or DD-MMM-YYYY
                dates = pd.to_datetime(expiries)
                # Return sorted list of strings.
                return sorted(expiries, key=lambda x: pd.to_datetime(x))
            except:
                return sorted(expiries)
                
        except Exception as e:
            self.log(f"Error reading expiries: {e}")
            return []

    def process(self, export_excel=True, export_csv=True, target_expiry=None):
        try:
            self.log(f"Opening ZIP: {os.path.basename(self.zip_path)}")
            
            # Load Daywise Data (IV, IVP, IVR)
            daywise_data = self.get_daywise_data()
            
            with zipfile.ZipFile(self.zip_path, 'r') as z:
                # Find the Bhav Copy CSV
                csv_filename = None
                for name in z.namelist():
                    if name.lower().endswith('.csv') and 'bhav' in name.lower():
                        csv_filename = name
                        break
                
                if not csv_filename:
                    raise FileNotFoundError("No valid Bhav Copy CSV found inside ZIP.")

                self.log(f"Extracting: {csv_filename}")
                with z.open(csv_filename) as f:
                    # Read CSV directly from memory
                    df = pd.read_csv(f)
            

            # Normalize columns to handle both Old and New formats
            col_map = {
                'FinInstrmTp': 'INSTRUMENT',
                'TckrSymb': 'SYMBOL',
                'XpryDt': 'EXPIRY_DT',
                'StrkPric': 'STRIKE_PR',
                'OptnTp': 'OPTION_TYP',
                'ClsPric': 'CLOSE'
            }
            # Rename if columns exist
            df.rename(columns=col_map, inplace=True)
            
            # Normalize column values for New Format
            # STF -> FUTSTK, STO -> OPTSTK
            if 'INSTRUMENT' in df.columns:
                 df['INSTRUMENT'] = df['INSTRUMENT'].replace({'STF': 'FUTSTK', 'STO': 'OPTSTK', 'IDF': 'FUTIDX', 'IDO': 'OPTIDX'})

            # Required columns validation
            required_cols = ['INSTRUMENT', 'SYMBOL', 'EXPIRY_DT', 'STRIKE_PR', 'OPTION_TYP', 'CLOSE']
            missing = [c for c in required_cols if c not in df.columns]
            if missing:
                raise ValueError(f"Missing columns in CSV: {missing}")

            # Filter for FUTSTK
            self.log("Filtering Futures data...")
            futures_df = df[df['INSTRUMENT'] == 'FUTSTK'].copy()
            
            if futures_df.empty:
                raise ValueError("No FUTSTK data found in file.")

            results = []
            
            unique_contracts = futures_df[['SYMBOL', 'EXPIRY_DT']].drop_duplicates()
            
            # Filter by Target Expiry if specified
            if target_expiry:
                self.log(f"Filtering for Expiry: {target_expiry}")
                unique_contracts = unique_contracts[unique_contracts['EXPIRY_DT'] == target_expiry]
            
            total_contracts = len(unique_contracts)
            self.log(f"Found {total_contracts} unique contracts. Processing...")

            # Pre-filter Options data to speed up lookups
            options_df = df[df['INSTRUMENT'] == 'OPTSTK'].copy()
            
            # Group options by Symbol and Expiry for faster access
            options_grouped = options_df.groupby(['SYMBOL', 'EXPIRY_DT'])

            count = 0
            skipped = 0

            for _, row in unique_contracts.iterrows():
                symbol = row['SYMBOL']
                expiry = row['EXPIRY_DT']
                
                # Get Future Close
                fut_row = futures_df[
                    (futures_df['SYMBOL'] == symbol) & 
                    (futures_df['EXPIRY_DT'] == expiry)
                ]
                
                if fut_row.empty:
                    continue
                    
                future_close = fut_row.iloc[0]['CLOSE']
                
                # Get Options Chain for this Symbol + Expiry
                if (symbol, expiry) not in options_grouped.groups:
                    skipped += 1
                    continue
                    
                chain = options_grouped.get_group((symbol, expiry))
                
                # Find Available Strikes
                available_strikes = chain['STRIKE_PR'].unique()
                
                if len(available_strikes) == 0:
                    skipped += 1
                    continue

                # Find ATM Strike (Nearest to Future Close)
                atm_strike = min(available_strikes, key=lambda x: abs(x - future_close))
                
                # Get CE and PE prices
                ce_row = chain[
                    (chain['STRIKE_PR'] == atm_strike) & 
                    (chain['OPTION_TYP'] == 'CE')
                ]
                
                pe_row = chain[
                    (chain['STRIKE_PR'] == atm_strike) & 
                    (chain['OPTION_TYP'] == 'PE')
                ]
                
                ce_price = ce_row.iloc[0]['CLOSE'] if not ce_row.empty else 0
                pe_price = pe_row.iloc[0]['CLOSE'] if not pe_row.empty else 0
                
                if ce_row.empty or pe_row.empty:
                    skipped += 1
                    continue

                # Get data from Daywise Report
                symbol_data = daywise_data.get(symbol, {})
                atm_iv = symbol_data.get('IV', 0)
                ivp = symbol_data.get('IVP', 0)
                ivr = symbol_data.get('IVR', 0)
                
                # Calculate Range, Upper, Lower
                # Formula: RANGE = CLOSE * IV% * SQRT(30/365)
                # IV is typically passed as whole number e.g. 23.5, so we use IV/100
                import math
                try:
                    iv_val = float(atm_iv)
                    range_val = future_close * (iv_val / 100.0) * math.sqrt(30.0 / 365.0)
                    upper_range = future_close + range_val
                    lowe_range = future_close - range_val
                except (ValueError, TypeError):
                    range_val = 0
                    upper_range = 0
                    lowe_range = 0

                # Calculate PROFIT %
                # PROFIT = RANGE - (CE price + PE price)
                # PROFIT % = (PROFIT / (CE price + PE price)) * 100
                total_premium = ce_price + pe_price
                if total_premium > 0:
                    profit = range_val - total_premium
                    profit_pct = (profit / total_premium) * 100
                else:
                    profit_pct = 0

                results.append({
                    'NAME': symbol,
                    'close': future_close,
                    'IV': atm_iv,
                    'IVP': ivp,
                    'IVR': ivr,
                    'LOWE RANGE': round(lowe_range, 2),
                    'RANGE': round(range_val, 2),
                    'UPPER RANGE': round(upper_range, 2),
                    'STRIKE': atm_strike,
                    'CE price': ce_price,
                    'PE price': pe_price,
                    'PROFIT %': round(profit_pct, 2)
                })
                
                count += 1
                if count % 100 == 0:
                     self.log(f"Processed {count} symbols...")

            if not results:
                raise ValueError("No valid ATM pairs found.")

            self.log(f"Processing complete. Valid: {len(results)}, Skipped: {skipped}")

            # Create Final DataFrame
            final_df = pd.DataFrame(results)
            
            # Reorder columns explicitly to match user request
            # NAME | close | IV | IVP | IVR | LOWE RANGE | RANGE | UPPER RANGE | STRIKE | CE price | PE price | PROFIT %
            desired_order = ['NAME', 'close', 'IV', 'IVP', 'IVR', 'LOWE RANGE', 'RANGE', 'UPPER RANGE', 'STRIKE', 'CE price', 'PE price', 'PROFIT %']
            final_df = final_df[desired_order]
            
            # Sort alphabetically by NAME
            final_df.sort_values(by='NAME', inplace=True)
            
            # Output filename handling
            date_str = "Output"
            
            import re
            base_name = os.path.basename(self.zip_path)
            match = re.search(r'\d{8}', base_name)
            if match:
                date_str = match.group(0)
            
            # Append Expiry to filename for clarity
            if target_expiry:
                # 2026-01-29 -> 29Jan2026
                try:
                    dt = pd.to_datetime(target_expiry)
                    exp_str = dt.strftime("%d%b%Y").upper()
                    date_str += f"_{exp_str}"
                except:
                    date_str += f"_{target_expiry}"
            
            base_dir = os.path.dirname(self.zip_path)
            files_created = []

            if export_excel:
                xlsx_path = os.path.join(base_dir, f"ATM_Options_{date_str}.xlsx")
                self.log(f"Saving Excel: {xlsx_path}")
                final_df.to_excel(xlsx_path, index=False)
                files_created.append(xlsx_path)

            if export_csv:
                csv_path = os.path.join(base_dir, f"ATM_Options_{date_str}.csv")
                self.log(f"Saving CSV: {csv_path}")
                final_df.to_csv(csv_path, index=False)
                files_created.append(csv_path)

            return files_created

        except Exception as e:
            self.log(f"Error: {str(e)}")
            raise e

