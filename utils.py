import base64
import streamlit as st

def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def set_branding():
    """
    Applies WHU branding to the Streamlit app.
    Includes CSS for colors, fonts, and sidebar logo.
    """
    
    # WHU Colors
    WHU_BLUE = "#2C4592"
    WHU_LIGHT_BLUE = "#808FBE"
    WHU_RED = "#E7331A"
    WHU_DARK_GRAY = "#515256"
    WHU_LIGHT_GRAY = "#EEEBEA"
    
    # Load Logo
    try:
        logo_base64 = get_base64_of_bin_file("assets/whu-logo-full.png")
        logo_html = f"""
            <div style="text-align: center; padding-bottom: 20px;">
                <img src="data:image/png;base64,{logo_base64}" alt="WHU Logo" style="width: 100%; max-width: 200px;">
            </div>
        """
        st.sidebar.markdown(logo_html, unsafe_allow_html=True)
    except FileNotFoundError:
        st.sidebar.warning("Logo not found in assets/whu-logo-full.png")

    # Custom CSS
    custom_css = f"""
        <style>
            /* Import Arial Font if not available (generic sans-serif fallback is usually fine) */
            @import url('https://fonts.googleapis.com/css2?family=Arial&display=swap');

            html, body, [class*="css"] {{
                font-family: 'Arial', sans-serif;
            }}

            /* Headers */
            h1, h2, h3, h4, h5, h6 {{
                color: {WHU_BLUE};
                font-family: 'Arial', sans-serif;
                font-weight: bold;
            }}

            /* Primary Button */
            div.stButton > button:first-child {{
                background-color: {WHU_BLUE};
                color: white;
                border: none;
                border-radius: 4px;
            }}
            div.stButton > button:first-child:hover {{
                background-color: {WHU_LIGHT_BLUE};
                color: white;
            }}

            /* Sidebar */
            [data-testid="stSidebar"] {{
                background-color: {WHU_LIGHT_GRAY};
            }}
            
            /* Links */
            a {{
                color: {WHU_RED};
            }}
            
            /* Expander Header */
            .streamlit-expanderHeader {{
                font-weight: bold;
                color: {WHU_BLUE};
            }}
            
        </style>
    """
    st.markdown(custom_css, unsafe_allow_html=True)
