import pandas as pd
import streamlit as st
import plotly.graph_objs as go
from datetime import datetime, timedelta
import os

ROOT_DIR  = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
activity_file_path = os.path.join(ROOT_DIR, 'data', "active_window_log.csv")
# Set the page configuration (Wide layout and Material design-like theme)
st.set_page_config(page_title="Window Activity Dashboard", layout="wide")


# Load the CSV data
@st.cache_data
def load_data(file_path):
    try:
        df = pd.read_csv(file_path)
        df['startTime'] = pd.to_datetime(df['startTime'])  # Ensure 'startTime' is a datetime
        return df
    except Exception as e:
        st.error(f"Error loading CSV file: {e}")
        return None


# Convert the total_time_spent to a readable format (H hours M min s sec)
def format_time_spent(seconds):
    # Convert seconds into a timedelta object
    time_spent = str(timedelta(seconds=int(seconds)))
    
    # timedelta outputs in the form 'H:MM:SS', so we need to manually format it for 'H hour M min s sec'
    hours, remainder = divmod(int(seconds), 3600)
    minutes, seconds = divmod(remainder, 60)
    
    # Construct the formatted string
    formatted_time = f"{hours} hour {minutes} min {seconds} sec" if hours > 0 else f"{minutes} min {seconds} sec"
    
    return formatted_time

# Define time filtering options
def filter_by_time(df, time_period):
    now = datetime.now()
    if time_period == "Current Hour":
        return df[df['startTime'] >= now.replace(minute=0, second=0, microsecond=0)]
    elif time_period == "Last Hour":
        return df[df['startTime'] >= now - timedelta(hours=1)]
    elif time_period == "Last 6 Hours":
        return df[df['startTime'] >= now - timedelta(hours=6)]
    elif time_period == "Today":
        return df[df['startTime'].dt.date == now.date()]
    elif time_period == "Yesterday":
        yesterday = now.date() - timedelta(days=1)
        return df[df['startTime'].dt.date == yesterday]
    elif time_period == "This Week":
        start_of_week = now - timedelta(days=now.weekday())
        return df[df['startTime'] >= start_of_week]
    else:
        return df

# Function to create the smoothed line chart
def create_smoothed_line_chart(df):
    fig = go.Figure()

    # Add the mouseClicks line with smoothing
    fig.add_trace(go.Scatter(
        x=df['startTime'], y=df['mouseClicks'],
        mode='lines+markers', name='Mouse Clicks',
        line=dict(color='blue', shape='spline'), marker=dict(size=6)
    ))

    # Add the keyPresses line with smoothing
    fig.add_trace(go.Scatter(
        x=df['startTime'], y=df['keyPresses'],
        mode='lines+markers', name='Key Presses',
        line=dict(color='green', shape='spline'), marker=dict(size=6)
    ))

    # Update layout for better visual clarity
    fig.update_layout(
        title="Mouse Clicks and Key Presses Over Time (Smoothed)",
        xaxis_title="Start Time",
        yaxis_title="Count",
        hovermode="x unified",
        template="plotly_white",
        margin=dict(l=0, r=0, t=50, b=0),
    )

    # Enable zooming
    fig.update_xaxes(rangeslider_visible=True)
    fig.update_yaxes(fixedrange=False)

    return fig

def change_name2_col(name):
    cols = ['activeWindow','activeProgram','activeTab','startTime','endTime','timeSpent','mouseClicks','keyPresses','task','goal']
    names = ['Active Window', 'Active Program', 'Active Tab', 'Start Time', 'End Time', 'Time Spent', 'Mouse Clicks', 'Key Presses', 'Task', 'Goal']
    return cols[names.index(name)]
# Function to create a summary table
def create_summary_table(df, group_by_columns):
    # Group the data by the selected group-by columns and sum the metrics
    col = change_name2_col(group_by_columns)
    df_summary = df.groupby(col).agg({
        'timeSpent': 'sum',
        'mouseClicks': 'sum',
        'keyPresses': 'sum'
    }).reset_index()

    df_summary['timeSpent'] = df_summary['timeSpent'].apply(format_time_spent)
    # Display the table in Streamlit
    st.subheader(f"Summary Table (Grouped by {group_by_columns})")
    st.dataframe(df_summary)

    return df_summary

def create_bar_chart(df, selected_dim,selected_metric):
    # Group the data by activeWindow and sum the selected metric
    col = change_name2_col(selected_dim)
    metrics = change_name2_col(selected_metric)

    df_grouped = df.groupby(col)[metrics].sum().reset_index()

    # Sort the DataFrame by the selected metric in descending order
    df_grouped = df_grouped.sort_values(by=metrics, ascending=False)

    # Define a list of colors for the bars
    colors = ['#33FF66', '#3366CC', '#330099', '#993399', '#FF0099', '#FFCC99', '#FFCC99', '#8d6e63']

    # Create a bar chart with activeWindow as x-axis and the summed selected metric as y-axis
    fig = go.Figure(
        data=[
            go.Bar(
                x=df_grouped[col],
                y=df_grouped[metrics],
                text=df_grouped[metrics],
                textposition='auto',
                name=selected_metric,
                marker_color=colors[:len(df_grouped)]  # Assign different colors to each bar
            )
        ]
    )

    # Update layout for better visual clarity
    fig.update_layout(
        title=f"Total {selected_metric} by {selected_dim} (Summed and Sorted)",
        xaxis_title=selected_dim,
        yaxis_title=selected_metric,
        template="plotly_white",
        margin=dict(l=0, r=0, t=50, b=0),
    )

    return fig

# Function to show top insights
def show_top_insights(df, group_by_columns):
    # Group the data and calculate sums for time spent, mouse clicks, and key presses
    col = change_name2_col(group_by_columns)
    time_spent_grouped = df.groupby(col)['timeSpent'].sum()
    mouse_clicks_grouped = df.groupby(col)['mouseClicks'].sum()
    key_presses_grouped = df.groupby(col)['keyPresses'].sum()

    # Check if the group-by result is empty
    if time_spent_grouped.empty:
        st.write("No data available for the selected grouping and time period.")
        return
    else:
        # Get the activeWindow with the most time spent and its corresponding value
        most_time_spent_window = time_spent_grouped.idxmax()
        most_time_spent_value = time_spent_grouped.max()

        most_time_spent_value = format_time_spent(most_time_spent_value)

        # Get the activeWindow with the most mouse clicks and its corresponding value
        most_mouse_clicks_window = mouse_clicks_grouped.idxmax()
        most_mouse_clicks_value = mouse_clicks_grouped.max()

        # Get the activeObject with the most key presses and its corresponding value
        most_key_presses_object = key_presses_grouped.idxmax()
        most_key_presses_value = key_presses_grouped.max()

        # Display the insights
        st.subheader(f"Key Insights ({group_by_columns})")
        st.write(f"Most time spent on: **{most_time_spent_window}** with **{most_time_spent_value} seconds**")
        st.write(f"Most mouse clicks on: **{most_mouse_clicks_window}** with **{most_mouse_clicks_value} clicks**")
        st.write(f"Most key presses on: **{most_key_presses_object}** with **{most_key_presses_value} key presses**")


# Sidebar for filtering options
st.sidebar.header("Filter Options")

# Load the data
file_path = activity_file_path
df = load_data(file_path)

if df is not None:
    # Dropdown for time period filtering
    time_period = st.sidebar.selectbox("Select Time Period", [
        "Current Hour", "Last Hour", "Last 6 Hours", "Today", "Yesterday", "This Week", "All Time"
    ])

    # Filtered DataFrame by time period (for both charts)
    df_filtered_by_time = filter_by_time(df, time_period)

    # Multiselect for filtering by activeWindow (applies only to line chart and metrics)
    active_window_list = df['activeWindow'].unique().tolist()
    # active_windows = st.sidebar.multiselect("Select Active Windows", active_window_list, default=active_window_list)

    # Checkbox options for summary table grouping
    st.sidebar.header("Active Windows Filter")
    active_windows = st.sidebar.multiselect(
        "Select Active Windows:", 
        options=active_window_list,
        default=active_window_list  # Default grouping is by activeWindow
    )

    # Filter the DataFrame for the line chart and metrics by activeWindow
    df_filtered_for_line_chart = df_filtered_by_time[df_filtered_by_time['activeWindow'].isin(active_windows)]

    # Compute Totals (applies only to the filtered line chart data)
    total_time_spent = df_filtered_for_line_chart['timeSpent'].sum()
    total_mouse_clicks = df_filtered_for_line_chart['mouseClicks'].sum()
    total_keypresses = df_filtered_for_line_chart['keyPresses'].sum()

    formatted_total_time_spent = format_time_spent(total_time_spent)

    # Display metrics as cards on the top
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Time Spent", formatted_total_time_spent)
    with col2:
        st.metric("Total Mouse Clicks", total_mouse_clicks)
    with col3:
        st.metric("Total Key Presses", total_keypresses)

    # --- Smoothed Line Chart Section (filtered by time and active windows) ---
    # st.subheader("Mouse Clicks and Key Presses Over Time (Smoothed)")
    line_chart = create_smoothed_line_chart(df_filtered_for_line_chart)
    st.plotly_chart(line_chart, use_container_width=True)

    # --- Bar Chart Section (filtered by time only, not by active windows) ---
    st.subheader("Bar Chart")

    # Select the metric for the bar chart
    selected_dim = st.selectbox("Select Dimension for Bar Chart", [
        'Active Window', 'Active Program', 'Active Tab', 'Task', 'Goal'
    ])
    # Select the metric for the bar chart
    selected_metric = st.selectbox("Select Metric for Bar Chart", [
        'Time Spent', 'Mouse Clicks', 'Key Presses'
    ])


    # Generate and display the bar chart (filtered only by time, not by active windows)
    bar_chart = create_bar_chart(df_filtered_by_time, selected_dim, selected_metric)
    st.plotly_chart(bar_chart, use_container_width=True)

    st.header("Summary Table")

    # Multiselect for summary table grouping (now in the main window)
    group_by_columns = st.selectbox(
        "Group Summary Table by:", 
        ['Active Window', 'Active Program', 'Active Tab', 'Task', 'Goal'],
    )

    # --- Summary Table Section (grouped by selected options) ---
    if group_by_columns:
        summary_table = create_summary_table(df_filtered_by_time, group_by_columns)
    # --- Top Insights Section ---
    show_top_insights(df_filtered_by_time, group_by_columns)
