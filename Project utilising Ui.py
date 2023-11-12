import tkinter as tk
import tkinter.font as tkFont
import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt
import warnings

warnings.filterwarnings("ignore")
plt.style.use('fivethirtyeight')

def apply_command():
    file_path = user_entry.get()
    start_year = int(start_year_entry.get())
    end_year = int(end_year_entry.get())
    product_category = category_entry.get()

    file_path = file_path.replace("\\", "/")

    try:
        df = pd.read_excel(file_path, engine='openpyxl')
    except Exception as e:
        print("Using 'xlrd' engine due to error:", e)
        df = pd.read_excel(file_path, engine='xlrd')

    df['Order Date'] = pd.to_datetime(df['Order Date'])
    mask = (df['Order Date'].dt.year >= start_year) & (df['Order Date'].dt.year <= end_year)
    filtered_data = df.loc[mask]

    product_data = filtered_data.loc[filtered_data['Category'] == product_category]  
    cols = ['Row ID', 'Order ID', 'Ship Date', 'Ship Mode', 'Customer ID']
    product_data.drop(cols, axis=1, inplace=True)
    product_data = product_data.set_index('Order Date')
    y = product_data['Sales'].resample('W').mean()
    y = y.fillna(method='ffill')
    return y

def show_sales_over_time():
    y = apply_command()
    plt.figure(figsize=(15, 6))
    plt.plot(y)
    plt.title('Sales over Time')
    plt.xlabel('Date')
    plt.ylabel('Sales')
    plt.show()

def show_sales_by_seasonal():
    y = apply_command()
    decomposition = sm.tsa.seasonal_decompose(y, model='additive')
    fig = decomposition.plot()
    plt.show()

def show_forecast():
    y = apply_command()
    mod = sm.tsa.statespace.SARIMAX(y,
                                    order=(1, 1, 1),
                                    seasonal_order=(1, 1, 0, 12),
                                    enforce_stationarity=False,
                                    enforce_invertibility=False)
    results = mod.fit()
    pred_uc = results.get_forecast(steps=100)
    pred_ci = pred_uc.conf_int()
    ax = y.plot(label='observed', figsize=(14, 7))
    pred_uc.predicted_mean.plot(ax=ax, label='Forecast')
    ax.fill_between(pred_ci.index, pred_ci.iloc[:, 0], pred_ci.iloc[:, 1], color='k', alpha=.25)
    ax.set_xlabel('Date')
    ax.set_ylabel('Furniture Sales')
    plt.legend()
    plt.show()

def show_model_diagnostics():
    y = apply_command()
    mod = sm.tsa.statespace.SARIMAX(y,
                                    order=(1, 1, 1),
                                    seasonal_order=(1, 1, 0, 12),
                                    enforce_stationarity=False,
                                    enforce_invertibility=False)
    results = mod.fit()
    results.plot_diagnostics(figsize=(16, 10))
    plt.show()

root = tk.Tk()
root.title("Sales Forecasting")
width = 600
height = 500
screenwidth = root.winfo_screenwidth()
screenheight = root.winfo_screenheight()
alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
root.geometry(alignstr)
root.resizable(width=False, height=False)

Label_1 = tk.Label(root)
ft = tkFont.Font(family='Times', size=23)
Label_1["font"] = ft
Label_1["fg"] = "#333333"
Label_1["justify"] = "center"
Label_1["text"] = "Your File Path"
Label_1.place(x=150, y=90, width=265, height=37)

user_entry = tk.Entry(root) 
user_entry.place(x=200, y=150, width=200, height=25)

start_year_label = tk.Label(root, text="Start Year:")  
start_year_label.place(x=150, y=200)
start_year_entry = tk.Entry(root)  
start_year_entry.place(x=220, y=200, width=50, height=25)

end_year_label = tk.Label(root, text="End Year:")  
end_year_label.place(x=150, y=250)
end_year_entry = tk.Entry(root) 
end_year_entry.place(x=220, y=250, width=50, height=25)

category_label = tk.Label(root, text="Product Category:")  
category_label.place(x=150, y=300)
category_entry = tk.Entry(root) 
category_entry.place(x=260, y=300, width=100, height=25)

sales_over_time_button = tk.Button(root, command=show_sales_over_time)
sales_over_time_button["bg"] = "#f0f0f0"
sales_over_time_button["cursor"] = "mouse"
sales_over_time_button["font"] = tkFont.Font(family='Times', size=10)
sales_over_time_button["fg"] = "#000000"
sales_over_time_button["justify"] = "center"
sales_over_time_button["text"] = "Sales over Time"
sales_over_time_button.place(x=150, y=350, width=120, height=25)

sales_by_seasonal_button = tk.Button(root, command=show_sales_by_seasonal)
sales_by_seasonal_button["bg"] = "#f0f0f0"
sales_by_seasonal_button["cursor"] = "mouse"
sales_by_seasonal_button["font"] = tkFont.Font(family='Times', size=10)
sales_by_seasonal_button["fg"] = "#000000"
sales_by_seasonal_button["justify"] = "center"
sales_by_seasonal_button["text"] = "Sales by Seasonal"
sales_by_seasonal_button.place(x=280, y=350, width=120, height=25)

forecast_button = tk.Button(root, command=show_forecast)
forecast_button["bg"] = "#f0f0f0"
forecast_button["cursor"] = "mouse"
forecast_button["font"] = tkFont.Font(family='Times', size=10)
forecast_button["fg"] = "#000000"
forecast_button["justify"] = "center"
forecast_button["text"] = "Forecast"
forecast_button.place(x=410, y=350, width=70, height=25)

model_diagnostics_button = tk.Button(root, command=show_model_diagnostics)
model_diagnostics_button["bg"] = "#f0f0f0"
model_diagnostics_button["cursor"] = "mouse"
model_diagnostics_button["font"] = tkFont.Font(family='Times', size=10)
model_diagnostics_button["fg"] = "#000000"
model_diagnostics_button["justify"] = "center"
model_diagnostics_button["text"] = "Model Diagnostics"
model_diagnostics_button.place(x=150, y=400, width=120, height=25)

root.mainloop()