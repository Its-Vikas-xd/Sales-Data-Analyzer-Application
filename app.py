import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
from matplotlib import cm
from matplotlib.colors import LinearSegmentedColormap

# Enhanced color scheme with more vibrant options
COLORS = {
    'background': '#2D2D2D',
    'foreground': '#FFFFFF',
    'accent1': '#00B4D8',  # Teal
    'accent2': '#FF6B6B',   # Coral
    'accent3': '#6AFF8B',   # Mint Green
    'accent4': '#FFD166',   # Yellow
    'accent5': '#A78BFA',   # Purple
    'secondary': '#4A4A4A',
    'text': '#E0E0E0',
    'chart_bg': '#1E1E1E'
}

# Custom color palettes
CATEGORY_PALETTE = [COLORS['accent1'], COLORS['accent2'], COLORS['accent3'], 
                   COLORS['accent4'], COLORS['accent5'], '#FF9F68', '#7BDCB5']
SEQUENTIAL_PALETTE = LinearSegmentedColormap.from_list("custom_sequential", 
                                                      [COLORS['chart_bg'], COLORS['accent1']])
HEATMAP_PALETTE = LinearSegmentedColormap.from_list("custom_heatmap", 
                                                   [COLORS['chart_bg'], COLORS['accent2']])

class DataAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Sales Analyzer")
        self.root.geometry("1000x800")
        self.file_path = None
        self.df = None
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_styles()
        
        # Create main container
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        ttk.Label(header_frame, text="SALES DATA ANALYZER", 
                 style='Header.TLabel', font=('Helvetica', 16, 'bold')).pack()
        ttk.Label(header_frame, text="Comprehensive Sales Analysis Dashboard", 
                 style='Subheader.TLabel').pack()
        
        # File selection section
        file_frame = ttk.Frame(self.main_frame)
        file_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(file_frame, text="Select Excel File:", style='Section.TLabel').pack(side=tk.LEFT, padx=5)
        self.file_entry = ttk.Entry(file_frame, width=50, style='Custom.TEntry')
        self.file_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="Browse", command=self.load_file, style='Accent.TButton').pack(side=tk.LEFT)
        
        # Analysis buttons
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(pady=15)
        
        buttons = [
            ("Data Overview", self.show_data_overview, COLORS['accent1']),
            ("Sales Visualizations", self.show_viz_options, COLORS['accent2']),
            ("Advanced Analysis", self.show_advanced_options, COLORS['accent3']),
            ("Exit", self.root.destroy, COLORS['secondary'])
        ]
        
        for text, command, color in buttons:
            btn = ttk.Button(button_frame, text=text, command=command, 
                           style=f'Custom.TButton')
            btn.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)
        
        # Output console
        console_frame = ttk.LabelFrame(self.main_frame, text="Analysis Output", style='Custom.TLabelframe')
        console_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.output_console = scrolledtext.ScrolledText(console_frame, height=15, wrap=tk.WORD,
                                                      bg=COLORS['secondary'], fg=COLORS['text'],
                                                      insertbackground=COLORS['text'],
                                                      font=('Consolas', 10))
        self.output_console.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.output_console.insert(tk.END, "Welcome to Sales Data Analyzer!\nPlease load an Excel file to begin...")
        
    def configure_styles(self):
        self.style.configure('.', background=COLORS['background'], foreground=COLORS['text'])
        self.style.configure('Custom.TEntry', fieldbackground=COLORS['secondary'],
                             foreground=COLORS['text'], bordercolor=COLORS['accent1'],
                             lightcolor=COLORS['accent1'], darkcolor=COLORS['accent1'])
        self.style.map('Custom.TButton',
                       foreground=[('active', COLORS['background']), ('!active', COLORS['text'])],
                       background=[('active', COLORS['accent1']), ('!active', COLORS['secondary'])],
                       bordercolor=[('active', COLORS['accent1'])])
        self.style.configure('Header.TLabel', font=('Helvetica', 16, 'bold'),
                             foreground=COLORS['accent1'])
        self.style.configure('Subheader.TLabel', font=('Helvetica', 10),
                             foreground=COLORS['accent2'])
        self.style.configure('Section.TLabel', font=('Helvetica', 10, 'bold'),
                             foreground=COLORS['accent1'])
        self.style.configure('Custom.TLabelframe', background=COLORS['background'],
                             foreground=COLORS['accent1'], bordercolor=COLORS['secondary'])
        self.style.configure('TLabelframe.Label', background=COLORS['background'],
                             foreground=COLORS['accent1'])

        # Configure matplotlib/seaborn styles
        plt.style.use('dark_background')
        sns.set_style("darkgrid", {
            'axes.facecolor': COLORS['chart_bg'],
            'grid.color': '#3A3A3A',
            'axes.edgecolor': COLORS['text'],
            'text.color': COLORS['text'],
            'axes.labelcolor': COLORS['text'],
            'xtick.color': COLORS['text'],
            'ytick.color': COLORS['text'],
        })
        plt.rcParams['figure.facecolor'] = COLORS['chart_bg']
        plt.rcParams['axes.titlecolor'] = COLORS['accent1']
        plt.rcParams['axes.titleweight'] = 'bold'
        plt.rcParams['axes.titlesize'] = 14

    def load_file(self):
        filetypes = (("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        filename = filedialog.askopenfilename(title="Open File", filetypes=filetypes)
        if filename:
            try:
                self.df = pd.read_excel(filename)
                self.file_path = filename
                self.file_entry.delete(0, tk.END)
                self.file_entry.insert(0, filename)
                self.output_console.delete(1.0, tk.END)
                self.output_console.insert(tk.END, f"File loaded successfully: {filename}\n")
                self.output_console.insert(tk.END, f"Dataset contains {self.df.shape[0]} rows and {self.df.shape[1]} columns\n")
                
                # Check for numeric columns and convert if needed
                self.convert_numeric_columns()
                
                messagebox.showinfo("Success", "File loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")
    
    def convert_numeric_columns(self):
        """Convert numeric columns to appropriate data types"""
        if self.df is None:
            return
            
        # Identify columns that should be numeric
        months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
        numeric_cols = months + ['Total Sales'] if 'Total Sales' in self.df.columns else months
        
        for col in numeric_cols:
            if col in self.df.columns and self.df[col].dtype != 'float64':
                try:
                    self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
                except Exception as e:
                    self.output_console.insert(tk.END, f"Warning: Could not convert {col} to numeric: {str(e)}\n")

    def show_error(self, msg):
        messagebox.showerror("Error", msg)

    def show_data_overview(self):
        if self.df is None:
            self.show_error("Please load a file first.")
            return
        
        self.output_console.delete(1.0, tk.END)
        self.output_console.insert(tk.END, "DATA OVERVIEW\n")
        self.output_console.insert(tk.END, "="*50 + "\n")
        
        # Basic info
        self.output_console.insert(tk.END, f"Dataset Dimensions: {self.df.shape[0]} rows x {self.df.shape[1]} columns\n\n")
        
        # Column data types
        self.output_console.insert(tk.END, "COLUMN DATA TYPES:\n")
        dtype_info = self.df.dtypes.to_string()
        self.output_console.insert(tk.END, dtype_info + "\n\n")
        
        # Descriptive statistics
        self.output_console.insert(tk.END, "DESCRIPTIVE STATISTICS:\n")
        overview = self.df.describe(include='all').to_string()
        self.output_console.insert(tk.END, overview + "\n\n")
        
        # Missing values
        self.output_console.insert(tk.END, "MISSING VALUES:\n")
        missing = self.df.isnull().sum().to_string()
        self.output_console.insert(tk.END, missing + "\n")

    def show_viz_options(self):
        if self.df is None:
            self.show_error("Please load a file first.")
            return
            
        viz_window = tk.Toplevel(self.root)
        viz_window.title("Select Visualization")
        viz_window.geometry("500x500")
        viz_window.configure(bg=COLORS['background'])
        
        header = ttk.Label(viz_window, text="Basic Visualizations", 
                          style='Header.TLabel', font=('Helvetica', 14))
        header.pack(pady=10)

        options = [
            ("1. Monthly Sales Trend", lambda: self.generate_chart(1)),
            ("2. Top Selling Products", lambda: self.generate_chart(2)),
            ("3. Monthly Sales Distribution", lambda: self.generate_chart(3)),
            ("4. Individual Product Trends", lambda: self.generate_chart(4)),
        ]

        for label, func in options:
            btn = ttk.Button(viz_window, text=label, command=func, style='Custom.TButton')
            btn.pack(padx=20, pady=10, fill=tk.X)
            
        ttk.Separator(viz_window, orient='horizontal').pack(fill=tk.X, padx=20, pady=5)
        
        close_btn = ttk.Button(viz_window, text="Close", command=viz_window.destroy,
                              style='Accent.TButton')
        close_btn.pack(pady=10, padx=20, fill=tk.X)

    def show_advanced_options(self):
        if self.df is None:
            self.show_error("Please load a file first.")
            return
            
        adv_window = tk.Toplevel(self.root)
        adv_window.title("Advanced Analysis")
        adv_window.geometry("500x500")
        adv_window.configure(bg=COLORS['background'])
        
        header = ttk.Label(adv_window, text="Advanced Visualizations", 
                          style='Header.TLabel', font=('Helvetica', 14))
        header.pack(pady=10)

        options = [
            ("5. Sales Heatmap by Product", lambda: self.generate_chart(5)),
            ("6. Product Sales Composition", lambda: self.generate_chart(6)),
            ("7. Monthly Sales Comparison", lambda: self.generate_chart(7)),
            ("8. Sales Distribution by Quarter", lambda: self.generate_chart(8)),
        ]

        for label, func in options:
            btn = ttk.Button(adv_window, text=label, command=func, style='Custom.TButton')
            btn.pack(padx=20, pady=10, fill=tk.X)
            
        ttk.Separator(adv_window, orient='horizontal').pack(fill=tk.X, padx=20, pady=5)
        
        close_btn = ttk.Button(adv_window, text="Close", command=adv_window.destroy,
                              style='Accent.TButton')
        close_btn.pack(pady=10, padx=20, fill=tk.X)

    def generate_chart(self, chart_choice):
        try:
            months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
            
            # Ensure months columns are numeric
            for col in months:
                if col in self.df.columns:
                    self.df[col] = pd.to_numeric(self.df[col], errors='coerce').fillna(0)
            
            # Create total sales column if it doesn't exist
            if 'Total Sales' not in self.df.columns:
                self.df['Total Sales'] = self.df[months].sum(axis=1)
            
            # Create a new figure with appropriate size
            if chart_choice in [4, 5, 7]:
                fig = plt.figure(figsize=(12, 8), facecolor=COLORS['chart_bg'])
            else:
                fig = plt.figure(figsize=(10, 6), facecolor=COLORS['chart_bg'])
                
            ax = fig.add_subplot(111)
            ax.set_facecolor(COLORS['chart_bg'])

            if chart_choice == 1:
                # Monthly Sales Trend (with area fill)
                monthly_totals = self.df[months].sum()
                ax.plot(monthly_totals.index, monthly_totals.values, 
                       marker='o', markersize=8, color=COLORS['accent1'], 
                       linewidth=2.5, alpha=0.9)
                ax.fill_between(monthly_totals.index, monthly_totals.values, 
                               color=COLORS['accent1'], alpha=0.2)
                ax.set_title("Monthly Sales Trend (All Products)", fontweight='bold')
                ax.set_ylabel("Sales Amount", fontweight='bold')
                ax.grid(True, alpha=0.2)

            elif chart_choice == 2:
                # Top Selling Products (horizontal bar)
                top_products = self.df.nlargest(8, 'Total Sales')
                sorted_products = top_products.sort_values('Total Sales', ascending=True)
                colors = CATEGORY_PALETTE[:len(sorted_products)]
                
                bars = ax.barh(sorted_products['Electrical Items'], sorted_products['Total Sales'], 
                              color=colors, alpha=0.9)
                
                # Add value labels
                for bar in bars:
                    width = bar.get_width()
                    ax.text(width * 1.01, bar.get_y() + bar.get_height()/2, 
                           f'${width:,.0f}', 
                           ha='left', va='center', color=COLORS['text'])
                
                ax.set_title("Top Selling Products (Annual Total)", fontweight='bold')
                ax.set_xlabel("Total Sales", fontweight='bold')
                ax.grid(True, alpha=0.2, axis='x')

            elif chart_choice == 3:
                # Monthly Sales Distribution (violin plot)
                melted_data = self.df[months].melt(var_name='Month', value_name='Sales')
                
                # Create a list of colors for the violin plot
                palette_list = [SEQUENTIAL_PALETTE(i) for i in np.linspace(0, 1, 12)]
                
                # Fixed violin plot with correct palette format
                sns.violinplot(x='Month', y='Sales', data=melted_data, 
                              palette=palette_list,
                              inner="quartile", ax=ax)
                ax.set_title("Monthly Sales Distribution", fontweight='bold')
                ax.set_ylabel("Sales Amount", fontweight='bold')
                ax.set_xlabel("Month", fontweight='bold')
                ax.grid(True, alpha=0.2)

            elif chart_choice == 4:
                # Individual Product Trends (grid of charts)
                fig.clf()
                top_5 = self.df.nlargest(6, 'Total Sales')
                fig, axes = plt.subplots(2, 3, figsize=(15, 10), facecolor=COLORS['chart_bg'])
                axes = axes.flatten()
                
                for idx, (_, row) in enumerate(top_5.iterrows()):
                    # Convert month values to float array
                    sales_values = np.array(row[months].values, dtype=float)
                    
                    # Create numerical indices for x-axis
                    x_indices = np.arange(len(months))
                    
                    axes[idx].plot(x_indices, sales_values, marker='o', markersize=5, 
                                 color=CATEGORY_PALETTE[idx], linewidth=2, alpha=0.9)
                    axes[idx].fill_between(x_indices, sales_values, 
                                         color=CATEGORY_PALETTE[idx], alpha=0.2)
                    axes[idx].set_title(row['Electrical Items'], fontsize=10, fontweight='bold')
                    axes[idx].set_xticks(x_indices)
                    axes[idx].set_xticklabels(months, rotation=45, ha='right')
                    axes[idx].set_facecolor(COLORS['chart_bg'])
                    axes[idx].grid(True, alpha=0.2)
                
                fig.suptitle("Top Products Monthly Sales Trends", fontsize=14, fontweight='bold')
                fig.tight_layout(rect=[0, 0, 1, 0.96])

            elif chart_choice == 5:
                # Sales Heatmap by Product (top products)
                fig.clf()
                top_products = self.df.nlargest(10, 'Total Sales')
                heatmap_data = top_products.set_index('Electrical Items')[months]
                
                fig, ax = plt.subplots(figsize=(12, 8), facecolor=COLORS['chart_bg'])
                sns.heatmap(heatmap_data, cmap=HEATMAP_PALETTE, annot=True, fmt=".0f", 
                           linewidths=0.5, linecolor=COLORS['secondary'], 
                           cbar_kws={'label': 'Sales Amount'}, ax=ax)
                ax.set_title("Monthly Sales Heatmap (Top Products)", fontweight='bold')
                ax.set_facecolor(COLORS['chart_bg'])
                plt.xticks(rotation=45)
                plt.tight_layout()

            elif chart_choice == 6:
                # Product Sales Composition (pie chart)
                top_products = self.df.nlargest(6, 'Total Sales')
                other_sales = self.df['Total Sales'].sum() - top_products['Total Sales'].sum()
                
                # Create a new row for 'Other' using concat instead of append
                other_row = pd.DataFrame([{'Electrical Items': 'Other', 'Total Sales': other_sales}])
                top_products = pd.concat([top_products, other_row], ignore_index=True)
                
                explode = [0.05] * len(top_products)
                colors = CATEGORY_PALETTE[:len(top_products)]
                
                wedges, texts, autotexts = ax.pie(
                    top_products['Total Sales'], 
                    labels=top_products['Electrical Items'], 
                    autopct='%1.1f%%', 
                    startangle=90, 
                    colors=colors, 
                    explode=explode, 
                    shadow=True,
                    textprops={'color': COLORS['text'], 'fontweight': 'bold'}
                )
                
                # Make percentages white and bold
                for autotext in autotexts:
                    autotext.set_color('white')
                    autotext.set_fontweight('bold')
                
                ax.set_title("Product Sales Composition", fontweight='bold')
                ax.axis('equal')  # Equal aspect ratio ensures pie is drawn as circle

            elif chart_choice == 7:
                # Monthly Sales Comparison (stacked area chart)
                top_products = self.df.nlargest(5, 'Total Sales')
                monthly_top = top_products[months].T
                monthly_top.columns = top_products['Electrical Items']
                
                # Create a stacked area chart
                ax.stackplot(monthly_top.index, monthly_top.T, 
                            labels=monthly_top.columns, 
                            colors=CATEGORY_PALETTE, alpha=0.8)
                
                ax.set_title("Monthly Sales Comparison (Top Products)", fontweight='bold')
                ax.set_ylabel("Sales Amount", fontweight='bold')
                ax.legend(loc='upper left')
                ax.grid(True, alpha=0.2)
                plt.xticks(rotation=45)

            elif chart_choice == 8:
                # Sales Distribution by Quarter
                self.df['Q1'] = self.df[['JAN', 'FEB', 'MAR']].sum(axis=1)
                self.df['Q2'] = self.df[['APR', 'MAY', 'JUN']].sum(axis=1)
                self.df['Q3'] = self.df[['JUL', 'AUG', 'SEP']].sum(axis=1)
                self.df['Q4'] = self.df[['OCT', 'NOV', 'DEC']].sum(axis=1)
                
                quarterly = self.df[['Q1', 'Q2', 'Q3', 'Q4']].sum()
                
                # Create a radial bar chart
                fig.clf()
                fig = plt.figure(figsize=(10, 8), facecolor=COLORS['chart_bg'])
                ax = fig.add_subplot(111, polar=True)
                
                # Compute angles
                N = len(quarterly)
                angles = [n / float(N) * 2 * np.pi for n in range(N)]
                angles += angles[:1]  # Close the circle
                
                # Prepare data
                values = quarterly.values.tolist()
                values += values[:1]
                
                # Plot the data
                ax.plot(angles, values, color=COLORS['accent1'], linewidth=2, marker='o', markersize=8)
                ax.fill(angles, values, color=COLORS['accent1'], alpha=0.2)
                
                # Add labels
                ax.set_xticks(angles[:-1])
                ax.set_xticklabels(quarterly.index, color=COLORS['text'], fontsize=10)
                ax.set_title("Quarterly Sales Distribution", fontweight='bold', pad=20)
                ax.grid(True, alpha=0.3)

            if chart_choice != 4:  # Subplots already handled in chart 4
                plt.tight_layout()

            chart_window = tk.Toplevel(self.root)
            chart_window.title("Analysis Result")
            chart_window.configure(bg=COLORS['background'])
            chart_window.geometry("1100x800")

            canvas = FigureCanvasTkAgg(fig, master=chart_window)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Add export button
            export_frame = ttk.Frame(chart_window)
            export_frame.pack(fill=tk.X, padx=10, pady=5)
            ttk.Button(export_frame, text="Export as PNG", 
                      command=lambda: self.export_figure(fig), 
                      style='Accent.TButton').pack(side=tk.RIGHT)

        except Exception as e:
            self.show_error(f"Chart Error: {str(e)}")
            import traceback
            traceback.print_exc()

    def export_figure(self, fig):
        filetypes = [('PNG Image', '*.png'), ('All Files', '*.*')]
        filename = filedialog.asksaveasfilename(filetypes=filetypes, defaultextension=".png")
        if filename:
            fig.savefig(filename, dpi=300, facecolor=COLORS['chart_bg'], bbox_inches='tight')
            messagebox.showinfo("Success", f"Chart saved successfully as:\n{filename}")

# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalysisApp(root)
    root.mainloop()
