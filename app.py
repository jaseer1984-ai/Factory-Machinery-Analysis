import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import os
import sys
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class FactoryAnalyzer:
    def __init__(self, excel_path=None):
        if excel_path is None:
            self.excel_path = Path("C:/Users/User/OneDrive/Documents/Factory_Project.xlsx")
        else:
            self.excel_path = Path(excel_path)
        
        self.output_path = Path("Investor_Project_Report.xlsx")
        self.charts_dir = Path("charts")
        self.charts_dir.mkdir(exist_ok=True)
        
        # Scenario efficiency assumptions
        self.scenario_efficiencies = {
            8: 0.955,   # Current efficiency from original data
            12: 0.80,
            16: 0.78, 
            20: 0.76,
            24: 0.75
        }
        
        self.load_data()
        
    def load_data(self):
        """Load all Excel sheets"""
        try:
            self.machines_df = pd.read_excel(self.excel_path, sheet_name='Machine Details')
            self.operating_df = pd.read_excel(self.excel_path, sheet_name='Operating Data')
            self.efficiency_df = pd.read_excel(self.excel_path, sheet_name='Efficiency & Losses')
            self.financial_df = pd.read_excel(self.excel_path, sheet_name='Financial Inputs')
            
            # Optional sales data
            try:
                self.sales_df = pd.read_excel(self.excel_path, sheet_name='Sales Report')
                self.has_sales_data = True
            except:
                self.sales_df = None
                self.has_sales_data = False
                print("No Sales Report sheet found - continuing without sales analysis")
                
            print(f"✅ Data loaded successfully from {self.excel_path}")
            
        except Exception as e:
            print(f"❌ Error loading Excel file: {e}")
            sys.exit(1)
    
    def calculate_production_metrics(self):
        """Calculate key production and financial metrics"""
        
        # Merge all dataframes on Machine ID
        self.df = self.machines_df.copy()
        
        # Add operating data
        if 'Machine_ID' in self.operating_df.columns:
            self.df = self.df.merge(self.operating_df, on='Machine_ID', how='left')
        else:
            # If no Machine_ID column, assume same order
            for col in self.operating_df.columns:
                if col != 'Machine_ID':
                    self.df[col] = self.operating_df[col].iloc[0] if len(self.operating_df) == 1 else self.operating_df[col]
        
        # Add efficiency data
        if 'Machine_ID' in self.efficiency_df.columns:
            self.df = self.df.merge(self.efficiency_df, on='Machine_ID', how='left')
        else:
            for col in self.efficiency_df.columns:
                if col != 'Machine_ID':
                    self.df[col] = self.efficiency_df[col].iloc[0] if len(self.efficiency_df) == 1 else self.efficiency_df[col]
        
        # Add financial data (typically same for all machines)
        financial_cols = ['raw_material_cost_per_1k', 'selling_price_per_1k', 'labor_cost_per_shift', 'electricity_cost_per_shift']
        for col in financial_cols:
            if col in self.financial_df.columns:
                self.df[col] = self.financial_df[col].iloc[0]
        
        # Calculate metrics
        self.df['theoretical_cups_per_day'] = (self.df['Capacity_CPM'] * 
                                             self.df['hours_per_day'] * 60)
        
        # Current efficiency calculation
        current_efficiency = self.df['actual_cups_per_day'].sum() / self.df['theoretical_cups_per_day'].sum()
        
        # Apply efficiency and downtime
        self.df['effective_hours'] = self.df['hours_per_day'] * (1 - self.df['downtime_hours'] / self.df['hours_per_day'])
        self.df['actual_production'] = (self.df['Capacity_CPM'] * 
                                       self.df['effective_hours'] * 60 * current_efficiency *
                                       (1 - self.df['wastage_percent'] / 100))
        
        # Financial calculations
        self.df['monthly_production'] = self.df['actual_production'] * self.df['working_days_per_month']
        
        # Revenue and costs per machine
        self.df['monthly_revenue'] = (self.df['monthly_production'] / 1000) * self.df['selling_price_per_1k']
        self.df['monthly_material_cost'] = (self.df['monthly_production'] / 1000) * self.df['raw_material_cost_per_1k']
        self.df['monthly_labor_cost'] = self.df['labor_cost_per_shift'] * self.df['working_days_per_month']
        self.df['monthly_electricity_cost'] = self.df['electricity_cost_per_shift'] * self.df['working_days_per_month']
        
        self.df['monthly_gross_margin'] = (self.df['monthly_revenue'] - 
                                         self.df['monthly_material_cost'] - 
                                         self.df['monthly_labor_cost'] - 
                                         self.df['monthly_electricity_cost'])
        
        # Calculate utilization
        self.df['utilization_percent'] = (self.df['actual_production'] / self.df['theoretical_cups_per_day']) * 100
        
        print("✅ Production metrics calculated")
    
    def calculate_scenarios(self):
        """Calculate different hour scenarios"""
        
        self.scenarios = {}
        base_capacity_per_minute = self.df['Capacity_CPM'].sum()
        base_working_days = self.df['working_days_per_month'].iloc[0]
        
        # Get financial parameters (assuming same for all scenarios)
        selling_price = self.df['selling_price_per_1k'].iloc[0]
        material_cost = self.df['raw_material_cost_per_1k'].iloc[0] 
        labor_cost = self.df['labor_cost_per_shift'].iloc[0]
        electricity_cost = self.df['electricity_cost_per_shift'].iloc[0]
        
        # Current scenario (baseline)
        current_hours = self.df['hours_per_day'].iloc[0]
        current_monthly_output = self.df['monthly_production'].sum()
        current_monthly_margin = self.df['monthly_gross_margin'].sum()
        
        for hours in [8, 12, 16, 20, 24]:
            efficiency = self.scenario_efficiencies[hours]
            
            # Calculate theoretical production
            daily_theoretical = base_capacity_per_minute * hours * 60
            monthly_theoretical = daily_theoretical * base_working_days
            
            # Apply efficiency and average wastage
            avg_wastage = self.df['wastage_percent'].mean()
            monthly_actual = monthly_theoretical * efficiency * (1 - avg_wastage / 100)
            
            # Calculate costs and margin
            monthly_revenue = (monthly_actual / 1000) * selling_price
            monthly_material_cost = (monthly_actual / 1000) * material_cost
            
            # Labor and electricity scale with hours/shifts
            shifts_per_day = hours / 8
            monthly_labor_cost = labor_cost * shifts_per_day * base_working_days
            monthly_electricity_cost = electricity_cost * shifts_per_day * base_working_days
            
            monthly_margin = (monthly_revenue - monthly_material_cost - 
                            monthly_labor_cost - monthly_electricity_cost)
            
            self.scenarios[hours] = {
                'hours': hours,
                'monthly_output': monthly_actual,
                'monthly_revenue': monthly_revenue,
                'monthly_margin': monthly_margin,
                'efficiency': efficiency,
                'margin_vs_current': monthly_margin - current_monthly_margin,
                'output_vs_current': monthly_actual - current_monthly_output
            }
        
        print("✅ Scenario calculations completed")
    
    def analyze_sales_gap(self):
        """Analyze sales vs production gap"""
        
        if not self.has_sales_data:
            self.sales_analysis = {
                'has_data': False,
                'message': 'No sales data available'
            }
            return
        
        # Calculate average monthly sales
        avg_monthly_sales = self.sales_df['monthly_sales_qty'].mean() if 'monthly_sales_qty' in self.sales_df.columns else 0
        current_monthly_production = self.df['monthly_production'].sum()
        
        gap = current_monthly_production - avg_monthly_sales
        
        # Calculate hours needed to match sales
        base_capacity = self.df['Capacity_CPM'].sum()
        working_days = self.df['working_days_per_month'].iloc[0]
        efficiency = self.scenario_efficiencies[8]  # Use current efficiency
        avg_wastage = self.df['wastage_percent'].mean()
        
        # Hours per day needed to match sales
        daily_sales_target = avg_monthly_sales / working_days
        theoretical_daily_capacity = base_capacity * 60  # per hour
        
        hours_needed = daily_sales_target / (theoretical_daily_capacity * efficiency * (1 - avg_wastage / 100))
        
        self.sales_analysis = {
            'has_data': True,
            'avg_monthly_sales': avg_monthly_sales,
            'current_monthly_production': current_monthly_production,
            'gap': gap,
            'gap_percentage': (gap / avg_monthly_sales) * 100 if avg_monthly_sales > 0 else 0,
            'hours_needed_for_sales': hours_needed,
            'is_overproducing': gap > 0
        }
        
        print("✅ Sales gap analysis completed")
    
    def create_charts(self):
        """Generate all charts"""
        
        plt.style.use('default')
        
        # Chart 1: Capacity vs Actual Production
        fig, ax = plt.subplots(figsize=(12, 6))
        x = range(len(self.df))
        ax.bar([i - 0.2 for i in x], self.df['theoretical_cups_per_day'], 
               width=0.4, label='Theoretical Capacity', alpha=0.7, color='lightblue')
        ax.bar([i + 0.2 for i in x], self.df['actual_production'], 
               width=0.4, label='Actual Production', alpha=0.7, color='darkblue')
        
        ax.set_xlabel('Machine ID')
        ax.set_ylabel('Cups per Day')
        ax.set_title('Theoretical vs Actual Production by Machine')
        ax.set_xticks(x)
        ax.set_xticklabels([f"M{i+1}" for i in x], rotation=45)
        ax.legend()
        plt.tight_layout()
        plt.savefig(self.charts_dir / 'capacity_vs_actual.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        # Chart 2: Utilization Histogram
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.hist(self.df['utilization_percent'], bins=15, alpha=0.7, color='green', edgecolor='black')
        ax.set_xlabel('Utilization %')
        ax.set_ylabel('Number of Machines')
        ax.set_title('Machine Utilization Distribution')
        ax.axvline(self.df['utilization_percent'].mean(), color='red', linestyle='--', 
                   label=f'Average: {self.df["utilization_percent"].mean():.1f}%')
        ax.legend()
        plt.tight_layout()
        plt.savefig(self.charts_dir / 'utilization_histogram.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        # Chart 3: Monthly Output by Scenario
        fig, ax = plt.subplots(figsize=(12, 6))
        hours = list(self.scenarios.keys())
        outputs = [self.scenarios[h]['monthly_output'] for h in hours]
        margins = [self.scenarios[h]['monthly_margin'] for h in hours]
        
        ax2 = ax.twinx()
        bars1 = ax.bar([f"{h}h" for h in hours], outputs, alpha=0.7, color='skyblue', label='Monthly Output')
        line1 = ax2.plot([f"{h}h" for h in hours], margins, 'ro-', label='Gross Margin (SAR)', linewidth=2, markersize=8)
        
        ax.set_xlabel('Operating Hours per Day')
        ax.set_ylabel('Monthly Output (Cups)', color='blue')
        ax2.set_ylabel('Monthly Gross Margin (SAR)', color='red')
        ax.set_title('Monthly Output & Margin by Operating Hours Scenario')
        
        # Add value labels on bars
        for bar, output in zip(bars1, outputs):
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{output:,.0f}', ha='center', va='bottom', fontsize=9)
        
        ax.legend(loc='upper left')
        ax2.legend(loc='upper right')
        plt.tight_layout()
        plt.savefig(self.charts_dir / 'scenario_analysis.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        # Chart 4: Sales vs Production (if data available)
        if self.has_sales_data:
            fig, ax = plt.subplots(figsize=(12, 6))
            
            # Create monthly data for visualization
            months = self.sales_df.index if 'month' not in self.sales_df.columns else self.sales_df['month']
            sales = self.sales_df['monthly_sales_qty'] if 'monthly_sales_qty' in self.sales_df.columns else [0] * len(months)
            production = [self.df['monthly_production'].sum()] * len(months)  # Constant production line
            
            ax.plot(months, sales, 'g-o', label='Monthly Sales', linewidth=2, markersize=6)
            ax.axhline(y=production[0], color='red', linestyle='--', 
                      label=f'Current Production Capacity: {production[0]:,.0f}', linewidth=2)
            
            ax.set_xlabel('Month')
            ax.set_ylabel('Cups')
            ax.set_title('Sales vs Production Capacity Trend')
            ax.legend()
            ax.grid(True, alpha=0.3)
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(self.charts_dir / 'sales_vs_production.png', dpi=300, bbox_inches='tight')
            plt.close()
        
        print("✅ Charts generated successfully")
    
    def generate_insights(self):
        """Generate narrative insights and recommendations"""
        
        # Calculate key metrics for insights
        total_machines = len(self.df)
        avg_utilization = self.df['utilization_percent'].mean()
        total_daily_output = self.df['actual_production'].sum()
        total_monthly_output = self.df['monthly_production'].sum()
        total_monthly_margin = self.df['monthly_gross_margin'].sum()
        
        # Top bottlenecks
        top_wastage = self.df.nlargest(3, 'wastage_percent')[['Machine_ID', 'wastage_percent']]
        top_downtime = self.df.nlargest(3, 'downtime_hours')[['Machine_ID', 'downtime_hours']]
        
        # Scenario comparison
        best_scenario_hours = max(self.scenarios.keys(), key=lambda x: self.scenarios[x]['monthly_margin'])
        best_scenario = self.scenarios[best_scenario_hours]
        
        insights_text = f"""
EXECUTIVE SUMMARY
================
• Factory operates {total_machines} machines with average utilization of {avg_utilization:.1f}%
• Current daily output: {total_daily_output:,.0f} cups
• Current monthly output: {total_monthly_output:,.0f} cups  
• Current monthly gross margin: {total_monthly_margin:,.0f} SAR

SCENARIO ANALYSIS
================
• Operating at {best_scenario_hours} hours/day could generate:
  - Monthly output: {best_scenario['monthly_output']:,.0f} cups ({best_scenario['output_vs_current']:+,.0f} vs current)
  - Monthly margin: {best_scenario['monthly_margin']:,.0f} SAR ({best_scenario['margin_vs_current']:+,.0f} vs current)

BOTTLENECK ANALYSIS
==================
Top 3 Wastage Issues:
{chr(10).join([f"• Machine {row['Machine_ID']}: {row['wastage_percent']:.1f}% rejection rate" for _, row in top_wastage.iterrows()])}

Top 3 Downtime Issues:  
{chr(10).join([f"• Machine {row['Machine_ID']}: {row['downtime_hours']:.1f} hours/day downtime" for _, row in top_downtime.iterrows()])}
"""

        if self.sales_analysis['has_data']:
            if self.sales_analysis['is_overproducing']:
                insights_text += f"""
SALES VS PRODUCTION ANALYSIS
===========================
• Average monthly sales: {self.sales_analysis['avg_monthly_sales']:,.0f} cups
• Current monthly production: {self.sales_analysis['current_monthly_production']:,.0f} cups
• OVERPRODUCTION: {self.sales_analysis['gap']:,.0f} cups ({self.sales_analysis['gap_percentage']:.0f}% excess)
• Factory could meet current sales demand with just {self.sales_analysis['hours_needed_for_sales']:.1f} hours/day operation
"""
            else:
                insights_text += f"""
SALES VS PRODUCTION ANALYSIS  
===========================
• Average monthly sales: {self.sales_analysis['avg_monthly_sales']:,.0f} cups
• Current monthly production: {self.sales_analysis['current_monthly_production']:,.0f} cups
• PRODUCTION SHORTFALL: {abs(self.sales_analysis['gap']):,.0f} cups
• Need {self.sales_analysis['hours_needed_for_sales']:.1f} hours/day operation to meet sales demand
"""

        recommendations = """
STRATEGIC RECOMMENDATIONS
========================
"""
        
        if self.sales_analysis['has_data'] and self.sales_analysis['is_overproducing']:
            recommendations += """
PRIORITY 1: MARKET DEVELOPMENT
• Focus on demand generation rather than capacity expansion
• Review pricing strategy to stimulate demand
• Develop new customer acquisition programs
• Consider market expansion to new regions/segments

PRIORITY 2: OPERATIONAL OPTIMIZATION  
• Target wastage reduction on high-rejection machines
• Each 1% wastage reduction directly improves gross margin
• Address recurring downtime through predictive maintenance
• Standardize operator training and preventive maintenance protocols

PRIORITY 3: CAPACITY OPTIMIZATION (Future)
• Only expand operating hours after establishing sustainable demand
• Consider 12-hour operations when sales reach 50% of current production
• Implement phased expansion: 12h → 16h → 20h based on market response
"""
        else:
            recommendations += """
PRIORITY 1: CAPACITY EXPANSION
• Increase operating hours to meet market demand
• Implement multi-shift operations for sustained production
• Focus on {best_scenario_hours}-hour operations for optimal margin

PRIORITY 2: EFFICIENCY IMPROVEMENTS
• Reduce wastage rates on underperforming machines  
• Minimize downtime through improved maintenance schedules
• Optimize shift changeover procedures

PRIORITY 3: MARKET DEVELOPMENT
• Continue expanding sales and marketing efforts
• Build strategic partnerships for demand growth
• Develop premium product lines for higher margins
"""

        self.insights_narrative = insights_text + recommendations
        
        print("✅ Insights and recommendations generated")
    
    def save_excel_report(self):
        """Save comprehensive Excel report"""
        
        with pd.ExcelWriter(self.output_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Format styles
            header_format = workbook.add_format({
                'bold': True, 'font_size': 12, 'bg_color': '#D7E4BC',
                'border': 1, 'align': 'center'
            })
            
            money_format = workbook.add_format({'num_format': '#,##0'})
            percent_format = workbook.add_format({'num_format': '0.0%'})
            
            # Sheet 1: Machine Summary (sorted by monthly margin)
            machine_summary = self.df[[
                'Machine_ID', 'Type', 'Capacity_CPM', 'actual_production', 
                'monthly_production', 'utilization_percent', 'wastage_percent',
                'downtime_hours', 'monthly_gross_margin'
            ]].sort_values('monthly_gross_margin', ascending=False)
            
            machine_summary.to_excel(writer, sheet_name='Machine Summary', index=False)
            worksheet1 = writer.sheets['Machine Summary']
            
            # Format headers
            for col_num, value in enumerate(machine_summary.columns.values):
                worksheet1.write(0, col_num, value, header_format)
            
            # Sheet 2: Scenario Analysis
            scenario_df = pd.DataFrame([
                {
                    'Operating Hours': f"{hours}h",
                    'Monthly Output (Cups)': int(data['monthly_output']),
                    'Monthly Revenue (SAR)': int(data['monthly_revenue']),
                    'Monthly Margin (SAR)': int(data['monthly_margin']),
                    'Efficiency %': f"{data['efficiency']*100:.1f}%",
                    'Margin vs Current': int(data['margin_vs_current']),
                    'Output vs Current': int(data['output_vs_current'])
                }
                for hours, data in self.scenarios.items()
            ])
            
            scenario_df.to_excel(writer, sheet_name='Scenario Analysis', index=False)
            worksheet2 = writer.sheets['Scenario Analysis']
            
            for col_num, value in enumerate(scenario_df.columns.values):
                worksheet2.write(0, col_num, value, header_format)
            
            # Sheet 3: Insights & Recommendations
            worksheet3 = workbook.add_worksheet('Insights & Recommendations')
            
            # Split insights into lines and write
            lines = self.insights_narrative.split('\n')
            for row, line in enumerate(lines):
                if any(header in line for header in ['EXECUTIVE', 'SCENARIO', 'BOTTLENECK', 'SALES', 'STRATEGIC', 'PRIORITY']):
                    worksheet3.write(row, 0, line, header_format)
                else:
                    worksheet3.write(row, 0, line)
            
            # Sheet 4: Sales Report (if available)
            if self.has_sales_data:
                self.sales_df.to_excel(writer, sheet_name='Sales Report', index=False)
                worksheet4 = writer.sheets['Sales Report']
                
                for col_num, value in enumerate(self.sales_df.columns.values):
                    worksheet4.write(0, col_num, value, header_format)
        
        print(f"✅ Excel report saved to: {self.output_path}")
    
    def print_console_summary(self):
        """Print executive summary to console"""
        
        print("\n" + "="*80)
        print("🏭 FACTORY OPERATIONS ANALYSIS - INVESTOR INSIGHTS")
        print("="*80)
        print(self.insights_narrative)
        print("="*80)
        print(f"📊 Charts saved to: {self.charts_dir}/")
        print(f"📈 Full report saved to: {self.output_path}")
        print("="*80)
    
    def run_analysis(self):
        """Execute complete analysis pipeline"""
        
        print("🚀 Starting Factory Analysis...")
        
        self.calculate_production_metrics()
        self.calculate_scenarios() 
        self.analyze_sales_gap()
        self.create_charts()
        self.generate_insights()
        self.save_excel_report()
        self.print_console_summary()
        
        print("\n✅ Analysis completed successfully!")
        
        return {
            'machine_data': self.df,
            'scenarios': self.scenarios,
            'sales_analysis': self.sales_analysis,
            'insights': self.insights_narrative
        }

def main():
    """Main execution function"""
    
    # Handle command line arguments or default path
    excel_path = sys.argv[1] if len(sys.argv) > 1 else None
    
    # Initialize and run analysis
    analyzer = FactoryAnalyzer(excel_path)
    results = analyzer.run_analysis()
    
    return results

if __name__ == "__main__":
    main()