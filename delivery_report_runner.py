"""
Delivery Report Runner
Main orchestrator that coordinates all modules to generate the complete report
"""

import os
import sys
import argparse
import logging
from datetime import datetime
from typing import Optional

# Import our modules
from input_parser import InputParser
from price_fetcher import PriceFetcher
from excel_writer import ExcelWriter

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class DeliveryReportGenerator:
    """Main class that orchestrates the entire process"""
    
    def __init__(self, input_file: str, mapping_file: str = "futures mapping.csv", usdinr_rate: float = 88.0):
        self.input_file = input_file
        self.mapping_file = mapping_file
        self.usdinr_rate = usdinr_rate
        self.parser = InputParser(mapping_file)
        self.price_fetcher = PriceFetcher()
    
    def generate_report(self, output_file: str = None):
        """Generate complete delivery report"""
        # Step 1: Parse input file
        logger.info(f"Parsing input file: {self.input_file}")
        positions = self.parser.parse_file(self.input_file)
        logger.info(f"Parsed {len(positions)} positions")
        
        if not positions:
            logger.error("No positions found in input file")
            return
        
        # Determine output filename based on format if not specified
        if not output_file:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            format_type = getattr(self.parser, 'format_type', 'UNKNOWN')
            
            if format_type in ['BOD', 'CONTRACT']:
                prefix = "GS_AURIGIN_DELIVERY"
            elif format_type == 'MS':
                prefix = "MS_WAFRA_DELIVERY"
            else:
                prefix = "DELIVERY_REPORT"
            
            output_file = f"{prefix}_{timestamp}.xlsx"
        
        logger.info("="*60)
        logger.info("Starting Delivery Report Generation")
        logger.info(f"Format Detected: {getattr(self.parser, 'format_type', 'UNKNOWN')}")
        logger.info(f"USDINR Rate: {self.usdinr_rate}")
        logger.info("="*60)
        
        # Step 2: Get unique symbols for price fetching
        symbol_map = {}
        for p in positions:
            symbol_map[p.underlying_ticker] = p.symbol
        
        symbols_to_fetch = list(set(p.symbol for p in positions))
        logger.info(f"Found {len(symbols_to_fetch)} unique symbols to fetch prices for")
        
        # Step 3: Fetch prices
        logger.info("Fetching prices from Yahoo Finance...")
        symbol_prices = self.price_fetcher.fetch_prices_for_symbols(symbols_to_fetch)
        
        # Map symbol prices to underlying tickers
        prices = {}
        for underlying, symbol in symbol_map.items():
            if symbol in symbol_prices:
                prices[underlying] = symbol_prices[symbol]
        
        logger.info(f"Mapped prices for {len(prices)} underlyings")
        
        # Step 4: Create Excel report
        logger.info("Creating Excel report...")
        writer = ExcelWriter(output_file, self.usdinr_rate)
        writer.create_report(positions, prices, self.parser.unmapped_symbols)
        
        logger.info("="*60)
        logger.info(f"Report generated successfully: {output_file}")
        logger.info("="*60)
        
        return output_file


def select_input_file():
    """Interactive file selection from current directory"""
    excel_files = []
    csv_files = []
    
    for file in os.listdir('.'):
        if file.endswith(('.xlsx', '.xls')):
            excel_files.append(file)
        elif file.endswith('.csv'):
            csv_files.append(file)
    
    all_files = excel_files + csv_files
    
    if not all_files:
        print("No Excel or CSV files found in current directory.")
        return None
    
    print("\n" + "="*60)
    print("SELECT INPUT FILE FOR DELIVERY CALCULATION")
    print("="*60)
    print("\nAvailable files in current directory:\n")
    
    file_index = 1
    file_map = {}
    
    if excel_files:
        print("Excel Files:")
        for file in sorted(excel_files):
            print(f"  [{file_index}] {file}")
            file_map[file_index] = file
            file_index += 1
    
    if csv_files:
        print("\nCSV Files:")
        for file in sorted(csv_files):
            print(f"  [{file_index}] {file}")
            file_map[file_index] = file
            file_index += 1
    
    print("\n" + "-"*60)
    
    while True:
        try:
            choice = input(f"\nEnter file number (1-{len(all_files)}) or 'q' to quit: ").strip()
            
            if choice.lower() == 'q':
                print("Exiting...")
                return None
            
            choice_num = int(choice)
            if 1 <= choice_num <= len(all_files):
                selected_file = file_map[choice_num]
                print(f"\nSelected: {selected_file}")
                return selected_file
            else:
                print(f"Please enter a number between 1 and {len(all_files)}")
        except ValueError:
            print("Invalid input. Please enter a number or 'q' to quit.")
        except KeyboardInterrupt:
            print("\n\nExiting...")
            return None


def main():
    parser = argparse.ArgumentParser(description='Generate Physical Delivery Report from Position File')
    parser.add_argument('input_file', nargs='?', help='Input position file (Excel or CSV)')
    parser.add_argument('--output', help='Output Excel file name')
    parser.add_argument('--usdinr', type=float, default=88.0, help='USDINR exchange rate (default: 88)')
    parser.add_argument('--interactive', '-i', action='store_true', help='Interactive mode for file selection')
    
    args = parser.parse_args()
    
    # Determine input file
    if args.input_file and not args.interactive:
        input_file = args.input_file
    else:
        input_file = select_input_file()
        if not input_file:
            sys.exit(0)
    
    # Verify input file exists
    if not os.path.exists(input_file):
        logger.error(f"Input file not found: {input_file}")
        sys.exit(1)
    
    # Fixed mapping file - no longer asking for permission
    mapping_file = 'futures mapping.csv'
    
    # Verify mapping file exists
    if not os.path.exists(mapping_file):
        logger.error(f"Mapping file not found: {mapping_file}")
        print("\nPlease ensure 'futures mapping.csv' is in the current directory.")
        sys.exit(1)
    
    # Get USDINR rate if interactive
    usdinr_rate = args.usdinr
    if args.interactive:
        usdinr_input = input(f"\nUSDINR exchange rate (default: 88): ").strip()
        if usdinr_input:
            try:
                usdinr_rate = float(usdinr_input)
            except ValueError:
                print("Invalid rate, using default: 88")
                usdinr_rate = 88.0
    
    # Output filename will be auto-generated based on format
    output_file = args.output  # Will be None if not specified, auto-generated in generator
    
    print("\n" + "="*60)
    print("STARTING DELIVERY REPORT GENERATION")
    print("="*60)
    print(f"Input File: {input_file}")
    print(f"Mapping File: {mapping_file}")
    print(f"USDINR Rate: {usdinr_rate}")
    print(f"Output File: Will be auto-generated based on format")
    print("="*60 + "\n")
    
    # Generate the report
    generator = DeliveryReportGenerator(input_file, mapping_file, usdinr_rate)
    output_filename = generator.generate_report(output_file)
    
    if output_filename:
        print(f"\nOutput saved as: {output_filename}")
    
    if args.interactive:
        input("\nPress Enter to exit...")
    
    sys.exit(0)


if __name__ == "__main__":
    # Check if no arguments provided, default to interactive mode
    if len(sys.argv) == 1:
        sys.argv.append('--interactive')
    main()