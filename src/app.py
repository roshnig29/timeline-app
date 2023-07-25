import dash
from dash import dcc, html
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output, State
import pandas as pd
import plotly.graph_objects as go
import base64
from plotly.subplots import make_subplots
from math import ceil
import webbrowser 
import threading

#Dash app and its elements
app = dash.Dash(__name__, external_stylesheets=[r'assets/styles.css'])
server = app.server

title = html.H1("Timeline graph")
upload = html.Div([
    dcc.Upload(
        id='upload-data',
        children=html.Div(['Drag and Drop or ',html.A('Select Files')]),
        multiple=False,
        className="upload-button",
        contents=None
    )
],className='button')
graph_height = html.Div([
    html.P("Height"),
    dcc.Input(
        id='height',
        className='textbox',
        type='number',
        value=900,  # Set the default value here
    )
], className='options')
graph_width = html.Div([
    html.P("Width"),
    dcc.Input(
        id='width',
        className='textbox',
        type='number',
        value=1500 # Set the default value here
    )
], className='options')
flag_pole = html.Div([
    html.P("Flag position"),
    dcc.Slider(
    id="pole",
    className="slider",
    max=10,
    min = -10,
    value= 2
)
], className='options')
graph_split = html.Div([
    html.P("Graph Split"),
    dcc.Slider(
    id="split",
    className="slider",
    max=1,
    min = 0,
    value= 0.65
)
], className='options')
arrow = html.Div([
    html.P("Arrow width"),
    dcc.Slider(
    id="arrow",
    className="slider",
    max=100,
    min = 5,
    value= 30
)
], className='options')
flag_height = html.Div([
    dcc.Dropdown(
        id='flag-dropdown',
        options=[],
        value=None,
        placeholder="Select flag name to adjust height",
    ),
    dcc.Slider(
    id="flag-height",
    className="slider",
    max=5,
    min =0.1,
    value= 0.5
)
], className='options',id='flag-options')

fonts = html.Div([
    html.P("Font Style"),
    dcc.Dropdown(
        id='font-style',
        options=[ "Arial", "Courier New" ,"Times New Roman"],
        value="Arial",
        placeholder="Select font",
    ),
        html.P("Font Size"),
        dcc.Input(
        id='font-size',
        className='textbox',
        type='number',
        value=18 # Set the default value here
    )
], className='options')

step1 = html.Div([
    html.H3("Step 1: Sheet 1"),
    html.Img(src='/assets/step1.png',),
    html.P("The first sheet contains the tasks with column headings 'Task','Start' and 'Finish'")
], className='steps')
step2 = html.Div([
    html.H3("Step 2: Sheet 2"),
    html.Img(src='/assets/step2.png',className='image-holder'),
    html.P("The second sheet contains the milestones with column headings 'Label','Date'",className='img-caption')
], className='steps')
step3 = html.Div([
    html.H3("Step 3: Save and upload"),
    html.Img(src='/assets/step3.png',),
    html.P("Save the file with project name and upload it to the website.")
], className='steps')
step4 = html.Div([
    html.H3("Step 4: Download as .png"),
    html.Img(src='/assets/step4.png',),
    html.P("Download the graph as an image file using the option on the top right.")
], className='steps')


progress_bar = html.Div(id='progress-container')
error_msg = html.Div(className='error')
button = html.Button('Create Graph', id='create-graph',n_clicks=0,className='button')
graph = html.Div(id='graph')

app.layout = dbc.Container([title,step1,step2,step3,step4,upload,progress_bar,graph_height,graph_width,flag_pole,graph_split,arrow,flag_height,fonts,graph],className='body')

#To provide a wrap on flag text
def add_line_break(text, break_length=13):
    words = text.split()
    lines = []
    current_line = ''
    
    for word in words:
        if len(current_line) + len(word) <= break_length:
            current_line += ' ' + word
        else:
            lines.append(current_line.strip())
            current_line = word
    
    lines.append(current_line.strip())
    
    return '<br>'.join(lines)

#Defines height of flags
def get_flag_positions(flags_labels,flag_num = None,pos=0.5):
    default = [0.5,1.0,0.2,1.2]
    if not y_value:
        for i in range(len(flags_labels)):
            y_value.append(default[i%4])
    elif flag_num!=None:
            try:
                i = flags_labels.index(flag_num)
                y_value[i]= pos
            except:
                return    
    return y_value

#Check for uploaded file
@app.callback(Output('progress-container', 'children'), [Input('upload-data', 'filename')])
def upload_progress(filename):
    if filename is not None and filename.endswith('.xlsx'):
        global y_value
        y_value = []
        return [html.H3('Uploaded file: {}'.format(filename), style={'color': 'green'})]
    elif filename is not None and not filename.endswith('.xlsx'):
        #raise ValueError("Only .xlsx files are allowed.")
        return html.Div([dbc.Alert('Error: Only .xlsx files are allowed!', color='danger',className='error')])
    else:
        return [html.H3('Please upload a file', style={'color': 'red'})]

#Plot the graph
@app.callback([Output('graph', 'children'),Output('flag-dropdown', 'options')], 
              [Input('upload-data', 'contents'),Input('height', 'value'), Input('width', 'value'), Input('pole', 'value'), 
               Input('split', 'value'), Input('arrow', 'value'), Input('flag-dropdown', 'value'), Input('flag-height', 'value'),
               Input('font-style', 'value'), Input('font-size', 'value')],
              State('upload-data','filename'))
def update_graph(contents,height,width,days,split,arrow,flag_num,flag_height,font_style,font_size,filename):
    if contents is not None:
        content_string = contents.split(',')[1]
        file = base64.b64decode(content_string)
    else:
        upload_progress(None)
    try:
        #Read excel files
        df_excel = pd.read_excel(file,sheet_name=0)
        df_excel.columns = df_excel.columns.str.lower()
        # Map the desired heading format to the column names
        column_mapping = {
        'task': 'Task',
        'start': 'Start',
        'finish': 'Finish',
        'date': 'date',
        'label':'label'
        }
        df_excel.columns = df_excel.columns.map(column_mapping.get) #Check for validity of columns
        flags_excel = pd.read_excel(file,sheet_name=1)
        flags_excel.columns = flags_excel.columns.str.lower()
        flags_excel.columns = flags_excel.columns.map(column_mapping.get)

        flag_labels = flags_excel['label'].tolist()

        df = df_excel.to_dict(orient='records')
        flags = flags_excel.to_dict(orient='records')
        
        for flag in flags:
            flag['label'] = add_line_break(flag['label'])


        start_dates = [item['Start'] for item in df]
        min_date = pd.to_datetime(min(start_dates))

        finish_dates = [item['Finish'] for item in df]
        max_date = pd.to_datetime(max(finish_dates))

        #xaxes range?
        range = max_date-min_date
        range_num=range.total_seconds()
        range_num = ceil(range_num/(3600*24*30)) 
        
        fig = make_subplots(rows=2,shared_xaxes=True) #Two subplots
        #Add the tasks
        for task in df[::-1]:
            #Start and end triangles
            fig.add_trace(
                go.Scatter(
                    x=[task['Start'], task['Finish']],
                    y=[task['Task'], task['Task']],
                    mode='lines+markers',
                    marker=dict(symbol='triangle-right', size=arrow+10),
                    name=task['Task'],
                    line=dict(width=arrow),
                ),row = 2,col = 1
            )
            #Join points and place text
            fig.add_annotation(
                    xref='x2',
                    yref='y2',
                    x=task["Start"],
                    y=task["Task"],
                    text=task["Task"],
                    showarrow=False,
                    xanchor='left',  # Anchor the text annotations to the right side of the plot area
                    xshift=+10,  # Adjust the x-shift as needed to position the text annotations properly
                    yanchor='middle',  # Anchor the text annotations to the middle of the y-axis category
                    font=dict(size=font_size, color='black'),
            )
        options = [{'label': flag_num, 'value': flag_num} for flag_num in flag_labels]
        y_value = get_flag_positions(flag_labels,flag_num,flag_height)

        tick_interval = pd.DateOffset(days=days)
        index = 0

        # Add the milestones
        for flag in flags:
            if('PQ' in flag['label']):
                color="cadetblue"
            elif("TP" in flag['label']):
                color = "coral"
            elif("today" in flag['label'].lower()):
                color = "darkblue"
            else:
                color = "turquoise"
            y1=y_value[index]
            date =  pd.to_datetime(flag['date'])
            #Add milestone flags
            fig.add_trace(
                go.Scatter(
                    x=[flag['date']+tick_interval], #Offset to position centroid to the right
                    y=[y1],  # Position the flag on the y-axis
                    mode='markers',
                    marker=dict(symbol='triangle-right', size=30,color=color),
                    name=flag['label'],
                ),row = 1,col = 1
            )
            #Add the lines
            fig.add_shape(
                type="line",
                x0=date, y0=0,  # Start of the line (x-axis)
                x1=date, y1=y1,  # End of the line (scatter point)
                line=dict(width=3,color=color)
            )
            #Add the text
            fig.add_annotation(
                x=flag['date'], y=y1+0.2,  # Position of the annotation
                text=flag['label'],
                showarrow=False,
                font=dict(size=font_size, color='black'),
            )
            index = (index+1)%len(y_value)

        #Layout changes
        fig.update_layout(
            font=dict(
                family=font_style,
                size=23,
                color="black",
            ),
            title = filename[:-5],
            height= height,
            width = width,
            xaxis2=dict(
            side="top", position=0.8,showgrid=True
            ),
            yaxis=dict(domain=[split,1],showticklabels=False, showgrid=False, zeroline=False),
            yaxis2=dict(domain=[0,split-0.05],showticklabels=False),
            showlegend = False,
        )
        return dcc.Graph(id='final-graph', figure=fig),options
    except KeyError as e:
        error_message = f"Error in format: {str(e)} is unknown. Please follow the requested format and upload file again."
        return html.Div(error_message,className='error'),[]
    except ValueError as e:
        error_message = f"Error in value: '{str(e)}'. Please check entered data and upload the file again."
        return html.Div(error_message,className='error'),[]
    except Exception as e:
        if contents is None:
            return None,[]
        error_message = f"Error in value: '{str(e)}'. Please check entered data and upload the file again."
        return html.Div(error_message,className='error'),[]


if __name__ == '__main__':
    #webbrowser.open(local_host_url)
    app.run_server(debug=True)