import streamlit as st
import pandas as pd
import plotly.express as px

# Load data
@st.cache_data
def load_data(path="results/design_results.xlsx"):
    return pd.read_excel(path)

df = load_data()

st.title("Design Parameter Plot Viewer")

numeric_cols = df.select_dtypes(include=["float", "int"]).columns.tolist()
if not numeric_cols:
    st.error("No numeric columns found.")
else:
    # Initialize session state keys with defaults
    if 'x_axis' not in st.session_state:
        st.session_state.x_axis = numeric_cols[0]

    if 'y_axis' not in st.session_state:
        if "stray_loss_pct" in numeric_cols:
            st.session_state.y_axis = "stray_loss_pct"
        else:
            default_y = next((col for col in numeric_cols if col != st.session_state.x_axis), numeric_cols[0])
            st.session_state.y_axis = default_y

    if 'color_param' not in st.session_state:
        color_options = [col for col in numeric_cols if col not in (st.session_state.x_axis, st.session_state.y_axis)]
        st.session_state.color_param = color_options[0] if color_options else st.session_state.x_axis

    # Select x-axis
    x_axis = st.selectbox(
        "Select X-axis",
        options=numeric_cols,
        index=numeric_cols.index(st.session_state.x_axis),
        key="x_axis"
    )

    # Select y-axis options excluding current x_axis
    y_axis_options = [col for col in numeric_cols if col != x_axis]
    if st.session_state.y_axis not in y_axis_options:
        st.session_state.y_axis = y_axis_options[0] if y_axis_options else numeric_cols[0]

    y_axis = st.selectbox(
        "Select Y-axis",
        options=y_axis_options,
        index=y_axis_options.index(st.session_state.y_axis),
        key="y_axis"
    )

    # Select color parameter options excluding x_axis and y_axis
    color_options = [col for col in numeric_cols if col not in (x_axis, y_axis)]
    if st.session_state.color_param not in color_options:
        st.session_state.color_param = color_options[0] if color_options else x_axis

    color_param = st.selectbox(
        "Select Color Parameter",
        options=color_options,
        index=color_options.index(st.session_state.color_param),
        key="color_param"
    )

    st.write(f"### Scatter Plot: {y_axis} vs {x_axis} (Color: {color_param})")

    # Plotly scatter with hover showing filename
    fig = px.scatter(
        df,
        x=x_axis,
        y=y_axis,
        color=color_param,
        hover_name="filename",
        color_continuous_scale="Viridis",
        labels={
            x_axis: x_axis.replace("_", " ").title(),
            y_axis: y_axis.replace("_", " ").title(),
            color_param: color_param.replace("_", " ").title(),
        },
        title=f"{y_axis.replace('_', ' ').title()} vs {x_axis.replace('_', ' ').title()} (Color by {color_param.replace('_', ' ').title()})"
    )

    fig.update_layout(
        width=1000,
        height=700,
        margin=dict(l=40, r=40, t=60, b=40),
        font=dict(
            family="Arial, sans-serif",
            size=14,
            color="black"
        ),
        xaxis=dict(
            title_font=dict(size=16, family="Arial, sans-serif", color="black"),
            tickfont=dict(size=14)
        ),
        yaxis=dict(
            title_font=dict(size=16, family="Arial, sans-serif", color="black"),
            tickfont=dict(size=14)
        ),
        title=dict(
            font=dict(size=18, family="Arial, sans-serif", color="black"),
            x=0.5,
            xanchor='center'
        ),
        coloraxis_colorbar=dict(
            title=dict(
                text=color_param.replace("_", " ").title(),
                side="right"  # default is "right", but this ensures vertical orientation
            )
        )
    )

    st.plotly_chart(fig, use_container_width=True)



# Launch the app:
# streamlit run app.py
# Open the browser to the link.