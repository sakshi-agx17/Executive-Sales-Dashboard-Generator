import pandas as pd
import numpy as np
from datetime import datetime
import xlsxwriter
import os
import warnings
warnings.filterwarnings('ignore')

class ExecutiveDashboard:
    def __init__(self):
        self.workbook = None
        self.data = None
        
    def load_global_superstore_data(self, filepath):
        """Load and process Global Superstore dataset from Kaggle"""
        try:
            print(f"Loading data from: {filepath}")
            for encoding in ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']:
                try:
                    df = pd.read_csv(filepath, encoding=encoding)
                    print(f"✓ Successfully loaded with {encoding} encoding")
                    break
                except UnicodeDecodeError:
                    continue
            else:
                raise Exception("Could not decode file with any standard encoding")
            
            print(f"✓ Loaded {len(df)} rows")
            print(f"✓ Columns: {df.columns.tolist()}")
            
            # Convert date columns
            date_col = None
            for col in ['Order Date', 'order_date', 'Date', 'OrderDate', 'Order_Date']:
                if col in df.columns:
                    date_col = col
                    break
            
            if date_col:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                df['Year'] = df[date_col].dt.year
                df['Month'] = df[date_col].dt.month
                df['MonthName'] = df[date_col].dt.strftime('%b')
                
                # years
                years = sorted(df['Year'].dropna().unique())
                print(f"✓ Years available: {years}")
                
                # last two years for comparison
                if len(years) >= 2:
                    year1, year2 = int(years[-2]), int(years[-1])
                else:
                    year1, year2 = int(years[0]), int(years[0])
                
                print(f"✓ Comparing {year1} vs {year2}")
                
                # Monthly sales aggregation
                monthly = df.groupby(['Year', 'Month', 'MonthName']).agg({
                    'Sales': 'sum',
                    'Profit': 'sum'
                }).reset_index()
                monthly = monthly.sort_values(['Year', 'Month'])
                
                sales_year1 = []
                sales_year2 = []
                
                months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
                
                for month_num in range(1, 13):
                    y1_data = monthly[(monthly['Year'] == year1) & (monthly['Month'] == month_num)]
                    y2_data = monthly[(monthly['Year'] == year2) & (monthly['Month'] == month_num)]
                    
                    sales_year1.append(float(y1_data['Sales'].sum()) if len(y1_data) > 0 else 0)
                    sales_year2.append(float(y2_data['Sales'].sum()) if len(y2_data) > 0 else 0)
                
            else:
                print("⚠ No date column found, using sample data")
                return False
            
            # Regional breakdown
            region_col = None
            for col in ['Region', 'region', 'Market', 'market', 'Country', 'country']:
                if col in df.columns:
                    region_col = col
                    break
            
            if region_col:
                regional = df.groupby(region_col).agg({
                    'Sales': 'sum',
                    'Profit': 'sum'
                }).reset_index()
                regional = regional.nlargest(10, 'Sales')
                
                regions = regional[region_col].tolist()
                regional_sales = regional['Sales'].tolist()
                regional_profit = regional['Profit'].tolist()
                
                regional_growth = [(p/s * 100) if s > 0 else 0 for s, p in zip(regional_sales, regional_profit)]
            else:
                regions = ['North', 'South', 'East', 'West', 'Central']
                regional_sales = [1000000, 800000, 1200000, 900000, 700000]
                regional_growth = [12, 8, 15, 10, 5]
            
            # Category breakdown
            category_col = None
            for col in ['Category', 'category', 'Product Category', 'Segment', 'Sub-Category']:
                if col in df.columns:
                    category_col = col
                    break
            
            if category_col:
                category = df.groupby(category_col)['Sales'].sum().reset_index()
                category = category.nlargest(10, 'Sales')
                
                categories = category[category_col].tolist()
                category_sales = category['Sales'].tolist()
            else:
                categories = ['Electronics', 'Furniture', 'Office Supplies', 'Technology', 'Accessories']
                category_sales = [2000000, 1500000, 1800000, 2200000, 1000000]
            
            # totals
            total_sales = float(df['Sales'].sum())
            total_profit = float(df['Profit'].sum()) if 'Profit' in df.columns else total_sales * 0.35
            
            # orders and customers
            order_col = None
            for col in ['Order ID', 'order_id', 'OrderID', 'Order_ID']:
                if col in df.columns:
                    order_col = col
                    break
            total_orders = int(df[order_col].nunique()) if order_col else len(df)
            
            customer_col = None
            for col in ['Customer ID', 'customer_id', 'CustomerID', 'Customer_ID']:
                if col in df.columns:
                    customer_col = col
                    break
            total_customers = int(df[customer_col].nunique()) if customer_col else int(total_orders * 0.3)
            
            # Targets
            sales_target = total_sales * 1.2
            profit_target = total_profit * 1.2
            
            # Update data dictionary
            self.data = {
                'months': months,
                'sales_year1': sales_year1,
                'sales_year2': sales_year2,
                'year1': year1,
                'year2': year2,
                'regions': regions,
                'regional_sales': regional_sales,
                'regional_growth': regional_growth,
                'categories': categories,
                'category_sales': category_sales,
                'total_sales': total_sales,
                'total_profit': total_profit,
                'total_customers': total_customers,
                'total_orders': total_orders,
                'sales_target': sales_target,
                'profit_target': profit_target
            }
            
            print(f"\n✓ Data processed successfully!")
            print(f"  Total Sales: ${total_sales:,.0f}")
            print(f"  Total Profit: ${total_profit:,.0f}")
            print(f"  Total Orders: {total_orders:,}")
            print(f"  Regions: {len(regions)}")
            print(f"  Categories: {len(categories)}")
            
            return True
            
        except Exception as e:
            print(f"✗ Error loading data: {e}")
            import traceback
            traceback.print_exc()
            print("  Using sample data instead...")
            return False
    
    def generate_sample_sales_data(self):
        """Generate realistic sample data"""
        print("Generating sample sales data...")
        
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        
        np.random.seed(42)
        sales_year1 = [180000 + np.random.randint(-20000, 30000) for _ in range(12)]
        sales_year2 = [200000 + np.random.randint(-20000, 35000) for _ in range(12)]
        
        self.data = {
            'months': months,
            'sales_year1': sales_year1,
            'sales_year2': sales_year2,
            'year1': 2021,
            'year2': 2022,
            'regions': ['North America', 'Europe', 'Asia Pacific', 'Latin America', 'Middle East'],
            'regional_sales': [3245000, 2890000, 4120000, 1560000, 980000],
            'regional_growth': [12.5, 8.3, 18.7, 15.2, 22.1],
            'categories': ['Technology', 'Furniture', 'Office Supplies'],
            'category_sales': [4200000, 3100000, 2800000],
            'total_sales': 12795000,
            'total_profit': 1467000,
            'total_customers': 1590,
            'total_orders': 51290,
            'sales_target': 15000000,
            'profit_target': 1800000
        }
        
        return self.data
    
    def create_dashboard(self, filename):
        """Create the complete visual dashboard"""
        self.workbook = xlsxwriter.Workbook(filename)
        
        # Create worksheet
        ws = self.workbook.add_worksheet('Executive Dashboard')
        ws.set_column('A:Z', 2.5)
        
        # Define formats
        title_format = self.workbook.add_format({
            'font_size': 24,
            'bold': True,
            'font_color': '#1F4E78',
            'align': 'left',
            'valign': 'vcenter'
        })
        
        subtitle_format = self.workbook.add_format({
            'font_size': 10,
            'italic': True,
            'font_color': '#7F8C8D',
            'align': 'left'
        })
        
        kpi_title_format = self.workbook.add_format({
            'font_size': 10,
            'bold': True,
            'font_color': '#1F4E78',
            'bg_color': '#F8F9FA',
            'align': 'left',
            'valign': 'top',
            'border': 1,
            'border_color': '#E0E0E0'
        })
        
        kpi_value_format = self.workbook.add_format({
            'font_size': 20,
            'bold': True,
            'font_color': '#2C3E50',
            'bg_color': '#F8F9FA',
            'align': 'left',
            'valign': 'vcenter',
            'border': 1,
            'border_color': '#E0E0E0'
        })
        
        kpi_percent_format = self.workbook.add_format({
            'font_size': 16,
            'bold': True,
            'font_color': '#1F4E78',
            'bg_color': '#F8F9FA',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'border_color': '#E0E0E0'
        })
        
        # Title
        ws.merge_range('B2:P2', f'GLOBAL SUPERSTORE DASHBOARD {self.data["year2"]}', title_format)
        ws.write('B3', 'Figures in USD', subtitle_format)
        
        # KPI Cards
        sales_pct = int((self.data['total_sales'] / self.data['sales_target']) * 100)
        profit_pct = int((self.data['total_profit'] / self.data['profit_target']) * 100)
        
        # Card 1: Total Sales
        ws.merge_range('B5:E5', 'Total Sales', kpi_title_format)
        ws.merge_range('B6:C7', f"${self.data['total_sales']/1000000:.2f}M", kpi_value_format)
        ws.merge_range('D6:E7', f'{sales_pct}%', kpi_percent_format)
        
        # Card 2: Total Profit
        ws.merge_range('G5:J5', 'Total Profit', kpi_title_format)
        ws.merge_range('G6:H7', f"${self.data['total_profit']/1000000:.2f}M", kpi_value_format)
        ws.merge_range('I6:J7', f'{profit_pct}%', kpi_percent_format)
        
        # Card 3: Total Orders
        ws.merge_range('L5:O5', 'Total Orders', kpi_title_format)
        ws.merge_range('L6:M7', f"{self.data['total_orders']:,}", kpi_value_format)
        ws.merge_range('N6:O7', '87%', kpi_percent_format)
        
        self.write_chart_data(ws)
        
        # Create charts
        self.create_line_chart(ws)
        self.create_bar_chart(ws)
        self.create_pie_chart(ws)
        self.create_radar_chart(ws)
        
        # Create data sheet
        self.create_data_sheet()
        
        self.workbook.close()
    
    def write_chart_data(self, ws):
        """Write data tables for charts"""
        # Sales trend data
        row = 49
        ws.write(row, 1, 'Month')
        ws.write(row, 2, str(self.data['year1']))
        ws.write(row, 3, str(self.data['year2']))
        
        for i, month in enumerate(self.data['months']):
            ws.write(row + i + 1, 1, month)
            ws.write(row + i + 1, 2, self.data['sales_year1'][i])
            ws.write(row + i + 1, 3, self.data['sales_year2'][i])
        
        # Regional data
        ws.write(row, 5, 'Region')
        ws.write(row, 6, 'Sales')
        for i, (region, sales) in enumerate(zip(self.data['regions'], self.data['regional_sales'])):
            ws.write(row + i + 1, 5, region)
            ws.write(row + i + 1, 6, sales)
        
        # Category data
        ws.write(row, 8, 'Category')
        ws.write(row, 9, 'Sales')
        for i, (cat, sales) in enumerate(zip(self.data['categories'], self.data['category_sales'])):
            ws.write(row + i + 1, 8, cat)
            ws.write(row + i + 1, 9, sales)
    
    def create_line_chart(self, ws):
        """Create sales trend line chart"""
        chart = self.workbook.add_chart({'type': 'line'})
        
        # Add series
        chart.add_series({
            'name': f'={ws.name}!$C$50',
            'categories': f'={ws.name}!$B$51:$B$62',
            'values': f'={ws.name}!$C$51:$C$62',
            'line': {'color': '#2E86AB', 'width': 2.5},
            'marker': {'type': 'circle', 'size': 7, 'fill': {'color': '#2E86AB'}}
        })
        
        chart.add_series({
            'name': f'={ws.name}!$D$50',
            'categories': f'={ws.name}!$B$51:$B$62',
            'values': f'={ws.name}!$D$51:$D$62',
            'line': {'color': '#A23B72', 'width': 2.5},
            'marker': {'type': 'circle', 'size': 7, 'fill': {'color': '#A23B72'}}
        })
        
        chart.set_title({'name': f'{self.data["year1"]}-{self.data["year2"]} Sales Trend', 'name_font': {'size': 14, 'bold': True}})
        chart.set_x_axis({'name': 'Month'})
        chart.set_y_axis({'name': 'Sales ($)', 'num_format': '$#,##0'})
        chart.set_legend({'position': 'bottom'})
        chart.set_size({'width': 480, 'height': 300})
        chart.set_style(12)
        
        ws.insert_chart('B10', chart)
    
    def create_bar_chart(self, ws):
        """Create regional performance bar chart"""
        chart = self.workbook.add_chart({'type': 'column'})
        
        num_regions = len(self.data['regions'])
        
        chart.add_series({
            'name': 'Sales by Region',
            'categories': f'={ws.name}!$F$51:$F${50+num_regions}',
            'values': f'={ws.name}!$G$51:$G${50+num_regions}',
            'fill': {'color': '#1F4E78'},
            'gap': 150
        })
        
        chart.set_title({'name': f'Sales by Region {self.data["year2"]}', 'name_font': {'size': 14, 'bold': True}})
        chart.set_x_axis({'name': 'Region'})
        chart.set_y_axis({'name': 'Sales ($)', 'num_format': '$#,##0'})
        chart.set_legend({'none': True})
        chart.set_size({'width': 400, 'height': 300})
        chart.set_style(11)
        
        ws.insert_chart('K10', chart)
    
    def create_pie_chart(self, ws):
        """Create category breakdown pie chart"""
        chart = self.workbook.add_chart({'type': 'pie'})
        
        num_categories = len(self.data['categories'])
        
        chart.add_series({
            'name': 'Sales by Category',
            'categories': f'={ws.name}!$I$51:$I${50+num_categories}',
            'values': f'={ws.name}!$J$51:$J${50+num_categories}',
            'data_labels': {'percentage': True, 'position': 'best_fit'}
        })
        
        chart.set_title({'name': 'Sales by Category', 'name_font': {'size': 14, 'bold': True}})
        chart.set_size({'width': 400, 'height': 300})
        chart.set_style(26)
        
        ws.insert_chart('K28', chart)
    
    def create_radar_chart(self, ws):
        """Create satisfaction radar chart - using column chart as fallback"""
        chart = self.workbook.add_chart({'type': 'radar', 'subtype': 'filled'})
        
        # Write satisfaction data
        ws.write(69, 11, 'Metric')
        ws.write(69, 12, 'Score')
        
        metrics = ['Quality', 'Delivery', 'Service', 'Value', 'Overall']
        scores = [86, 78, 65, 82, 88]
        
        for i, (metric, score) in enumerate(zip(metrics, scores)):
            ws.write(70 + i, 11, metric)
            ws.write(70 + i, 12, score)
        
        chart.add_series({
            'name': 'Customer Satisfaction',
            'categories': f'={ws.name}!$L$71:$L$75',
            'values': f'={ws.name}!$M$71:$M$75',
            'fill': {'color': '#1F4E78', 'transparency': 50},
            'line': {'color': '#1F4E78', 'width': 2}
        })
        
        chart.set_title({'name': 'Customer Satisfaction Metrics', 'name_font': {'size': 14, 'bold': True}})
        chart.set_size({'width': 400, 'height': 300})
        chart.set_style(26)
        
        ws.insert_chart('B28', chart)
    
    def create_data_sheet(self):
        """Create detailed data sheet"""
        ws = self.workbook.add_worksheet('Data Source')
        
        header_format = self.workbook.add_format({
            'bold': True,
            'bg_color': '#1F4E78',
            'font_color': 'white',
            'border': 1
        })
        
        # Monthly data
        ws.write('A1', 'Monthly Sales Data', self.workbook.add_format({'font_size': 14, 'bold': True}))
        
        headers = ['Month', 'Year', 'Sales ($)', 'Profit ($)']
        for col, header in enumerate(headers):
            ws.write(1, col, header, header_format)
        
        row = 2
        for i, month in enumerate(self.data['months']):
            ws.write(row, 0, month)
            ws.write(row, 1, self.data['year1'])
            ws.write(row, 2, self.data['sales_year1'][i])
            ws.write(row, 3, self.data['sales_year1'][i] * 0.35)
            row += 1
        
        for i, month in enumerate(self.data['months']):
            ws.write(row, 0, month)
            ws.write(row, 1, self.data['year2'])
            ws.write(row, 2, self.data['sales_year2'][i])
            ws.write(row, 3, self.data['sales_year2'][i] * 0.35)
            row += 1
        
        # Regional summary
        ws.write(row + 2, 0, 'Regional Performance', self.workbook.add_format({'font_size': 14, 'bold': True}))
        
        headers = ['Region', 'Sales ($)', 'Growth (%)']
        for col, header in enumerate(headers):
            ws.write(row + 3, col, header, header_format)
        
        for i, (region, sales, growth) in enumerate(zip(self.data['regions'], 
                                                         self.data['regional_sales'], 
                                                         self.data['regional_growth'])):
            ws.write(row + 4 + i, 0, region)
            ws.write(row + 4 + i, 1, sales)
            ws.write(row + 4 + i, 2, growth / 100, self.workbook.add_format({'num_format': '0.00%'}))
    
    def generate(self, filename=None):
        """Main generation function"""
        if filename is None:
            filename = '/Users/ignite/Downloads/Portfolio_Dashboard.xlsx'
        
        print("=" * 60)
        print("GENERATING EXECUTIVE VISUAL DASHBOARD")
        print("=" * 60)
        
        # Create dashboard
        self.create_dashboard(filename)
        
        print("\n" + "=" * 60)
        print("✓ DASHBOARD CREATED SUCCESSFULLY!")
        print("=" * 60)
        print(f"📁 Location: {filename}")
        print(f"💰 Total Sales: ${self.data['total_sales']:,.0f}")
        print(f"📈 Total Profit: ${self.data['total_profit']:,.0f}")
        print(f"📦 Total Orders: {self.data['total_orders']:,}")
        print(f"⏰ Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("\n📊 Dashboard includes:")
        print("   • Executive KPI cards with percentages")
        print("   • Sales trend line chart (year over year)")
        print("   • Regional performance bar chart")
        print("   • Category breakdown pie chart")
        print("   • Customer satisfaction radar chart")
        print("=" * 60)

def main():
    """Main execution"""
    dashboard = ExecutiveDashboard()
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # load data
    kaggle_file = os.path.join(script_dir, 'Global_Superstore2.csv')
    
    if os.path.exists(kaggle_file):
        print(f"Found data file in: {kaggle_file}")
        success = dashboard.load_global_superstore_data(kaggle_file)
        if not success:
            print("\nFalling back to sample data...")
            dashboard.generate_sample_sales_data()
    else:
        print(f"Data file not found: {kaggle_file}")
        print("Using sample data instead...")
        print("\n💡 TIP: Place 'Global_Superstore2.csv' in the same folder as this script")
        dashboard.generate_sample_sales_data()
    
    # Generate dashboard
    output_file = os.path.join(script_dir, 'Portfolio_Dashboard.xlsx')
    dashboard.generate(output_file)

if __name__ == "__main__":
    main()
