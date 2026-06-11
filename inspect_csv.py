import pandas as pd
import os

csv_path = r"C:\Users\hp\Downloads\invoice-reconciliation-thirdPartyInvoice (2).csv"
print(f"Checking CSV file: {csv_path}")
print(f"Exists: {os.path.exists(csv_path)}")

if os.path.exists(csv_path):
    try:
        df = pd.read_csv(csv_path)
        print(f"\nCSV Dimensions: {df.shape}")
        print(f"Columns in CSV: {list(df.columns)}")
        
        # Broadcaster column could be "Third Party Vendor Name" or "Channel Name"
        vendor_col = None
        for col in df.columns:
            if "Vendor" in col or "Broadcaster" in col or "Publisher" in col:
                vendor_col = col
                break
                
        channel_col = None
        for col in df.columns:
            if "Channel" in col:
                channel_col = col
                break
                
        if vendor_col:
            print(f"\nDistribution by Vendor ('{vendor_col}'):")
            print(df[vendor_col].value_counts())
        else:
            print("\nVendor column not found.")
            
        if channel_col:
            print(f"\nDistribution by Channel ('{channel_col}'):")
            print(df[channel_col].value_counts().head(20))
            
        # File Name distribution
        if "File Name" in df.columns:
            print(f"\nTotal unique PDF files present in thirdPartyInvoice: {df['File Name'].nunique()}")
            print("Top 10 files by row count:")
            print(df["File Name"].value_counts().head(10))
            
    except Exception as e:
        print(f"Error reading CSV: {e}")
