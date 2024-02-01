import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import openpyxl
import pandas as pd

# Load data from Excel
excel_dataframe = openpyxl.load_workbook("NewData.xlsx")
dataframe = excel_dataframe.active

# Convert data from Excel to pandas DataFrame
data = []
for row in range(1, dataframe.max_row + 1):
    _row = [dataframe.cell(row=row, column=1).value]
    for col in range(2, dataframe.max_column + 1):
        _row.append(dataframe.cell(row=row, column=col).value)
    data.append(_row)

headers = ["State"] + ["2011", "2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023"]
df = pd.DataFrame(data, columns=headers)

# Initialize Dash application
app = dash.Dash(__name__, url_base_pathname='/Security-Crime-Incidence/')


# App layout
app.layout = html.Div([
    html.H1("SECURITY CRIME INCIDENCE"),

    # Dropdown for selecting state(s)
    dcc.Dropdown(
        id='dropdown-state',
        options=[
            {'label': state, 'value': state} for state in df['State']
        ],
        multi=True,
        placeholder='Seleccione estado(s)'
    ),

    # Dropdown for selecting specific year(s)
    dcc.Dropdown(
        id='dropdown-year',
        options=[
            {'label': year, 'value': year} for year in df.columns[1:]
        ],
        multi=True,  # Allow multi-selection
        placeholder='Seleccione uno o más años'
    ),

    # Line chart
    dcc.Graph(
        id='line-chart',
    ),

    # Data Table for selected year(s)
    dash_table.DataTable(
        id='data-table',
        style_table={'overflowX': 'auto', 'width': '30%', 'float': 'left'},  # Ajusta el ancho y la posición de la tabla
    )
])

@app.callback(
    [Output('line-chart', 'figure'),
     Output('data-table', 'columns'),
     Output('data-table', 'data')],
    [Input('dropdown-state', 'value'),
     Input('dropdown-year', 'value')]
)
def update_chart(selected_states, selected_years):
    # Set default values for selected_states and selected_years if they are None
    selected_states = selected_states or []
    selected_years = selected_years or df.columns[1:].tolist()  # Use all available years if none are selected

    # Filter out None values in the selected_states list
    selected_states = [state for state in selected_states if state is not None]

    if selected_states:
        # Filter DataFrame by selected states
        filtered_df = df[df['State'].isin(selected_states)]

        # Create line chart
        traces = [
            {
                'x': filtered_df.columns[1:],
                'y': filtered_df.loc[filtered_df['State'] == state].iloc[0, 1:],
                'type': 'line',
                'name': state
            } for state in selected_states
        ]

        figure = {'data': traces, 'layout': {'title': 'Datos por Estado'}}

        # Define columns for DataTable based on selected years
        table_columns = [{'name': 'State', 'id': 'State'}] + [{'name': year, 'id': year} for year in selected_years]

        # Filter data for DataTable based on selected years
        table_data = [{'State': state} for state in selected_states]
        for year in selected_years:
            for state in selected_states:
                if year in filtered_df.columns:
                    table_data[selected_states.index(state)][year] = filtered_df.loc[filtered_df['State'] == state, year].values[0]
                else:
                    table_data[selected_states.index(state)][year] = None

        return figure, table_columns, table_data
    else:
        # If no state is selected, show empty chart and table
        return {'data': [], 'layout': {}}, [], []

if __name__ == '__main__':
    app.run_server(debug=True)
