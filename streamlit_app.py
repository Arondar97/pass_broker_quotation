import streamlit as st
import pandas as pd
from io import BytesIO
from quotation_creation import run_quotation_process

# Initialize session state for page navigation if it's not already set
if 'page' not in st.session_state:
    st.session_state.page = 'welcome'

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    
if 'user_role' not in st.session_state:
    st.session_state.user_role = None

# --- Excel File Configuration ---
EXCEL_USER_FILE = "users.xlsx"


def load_users_from_excel():
    """
    Loads user credentials and roles from a specified Excel file.
    The file must have columns: 'Utenza', 'Password', 'Ruolo'.
    """
    try:
        users_df = pd.read_excel(EXCEL_USER_FILE)
        return users_df
    except FileNotFoundError:
        st.error(f"Errore: Il file '{EXCEL_USER_FILE}' non è stato trovato. Si prega di creare il file con le colonne 'Utenza', 'Password' e 'Ruolo'.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Errore durante la lettura del file Excel: {e}")
        return pd.DataFrame()

def authenticate_user(username, password, users_df):
    """
    Checks if the provided username and password match a record in the DataFrame.
    Returns the user's role if authentication is successful, otherwise returns None.
    
    Args:
        username (str): The username entered by the user.
        password (str): The password entered by the user.
        users_df (pd.DataFrame): The DataFrame containing user credentials.
        
    Returns:
        str or None: The user's role ('Admin', 'Collaboratore', 'Esperto') or None.
    """
    if users_df.empty:
        return None
        
    # Find the row that matches both username and password
    user_row = users_df[
        (users_df['Utenza'] == username) &
        (users_df['Password'] == password)
    ]
    
    if not user_row.empty:
        # Return the 'Ruolo' from the matched row
        return user_row['Ruolo'].iloc[0]
    else:
        return None

def welcome_page():
    """
    Displays the initial welcome page with navigation buttons.
    """
    # Use columns to position the header to the left and higher
    col_empty, col_header, _ = st.columns([0.05, 3, 1]) 
    with col_header:
        st.markdown("# Pass Broker")

    # Use columns to center the content on the page
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # Title of the page
        st.title("Benvenuto!")
        st.write("---") # A horizontal line for visual separation

        # Button 1: "Cerca il preventivo"
        if st.button("Cerca il preventivo", use_container_width=True):
            st.success("Button 'Cerca il preventivo' was clicked!")
            # Future logic here for the 'search' page

        # Add a small vertical space between the buttons
        st.write("") 

        # Button 2: "Login"
        if st.button("Login", use_container_width=True):
            st.session_state.page = 'login'
            st.rerun()

def login_form_page():
    """
    Displays the login form with input fields and a button.
    """
    # Use columns to position the header to the left and higher
    col_empty, col_header, _ = st.columns([0.05, 3, 1])
    with col_header:
        st.markdown("# Pass Broker")

    # Use columns to center the login form
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("Login")
        st.write("---")
        
        # Input fields for username and password
        username = st.text_input("Utenza")
        password = st.text_input("Password", type="password")
        
        st.write("") 
        
        # "Accedi" (Login) button
        if st.button("Accedi", use_container_width=True):
            users_df = load_users_from_excel()
            user_role = authenticate_user(username, password, users_df)
            
            if user_role:
                st.success("Accesso eseguito con successo!")
                st.session_state.logged_in = True
                st.session_state.user_role = user_role
                st.session_state.page = 'dashboard'
                st.rerun()
            else:
                st.error("Credenziali non corrette. Riprova.")
        
        # Add a "back" button to return to the welcome page
        st.write("") 
        if st.button("Torna indietro", use_container_width=True):
            st.session_state.page = 'welcome'
            st.rerun()

def dashboard_page():
    """
    Displays the user dashboard with buttons based on their role.
    """
    # Security check: if not logged in, redirect to login page
    if not st.session_state.logged_in:
        st.warning("Per accedere a questa pagina, devi prima effettuare il login.")
        st.session_state.page = 'login'
        st.rerun()
        return

    # Use columns to position the header to the left and higher
    col_empty, col_header, _ = st.columns([0.05, 3, 1])
    with col_header:
        st.markdown("# Pass Broker")

    # Use columns to center the dashboard content
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title(f"Benvenuto, {st.session_state.user_role}!")
        st.write("---")
        
        st.write("Seleziona un'azione:")

        # Button 1: Calcola nuovo preventivo (Visible to 'Esperto')
        if st.session_state.user_role in ['Esperto', 'Admin']:
            if st.button("Calcola nuovo preventivo", use_container_width=True):
                st.session_state.page = 'quotation_calculation'
                st.rerun()

        # Button 2: Componi preventivo (Visible to all)
        if st.button("Componi preventivo", use_container_width=True):
            st.info("Funzionalità 'Componi preventivo' in sviluppo...")

        # Button 3: Consulta dati (Visible to 'Esperto')
        if st.session_state.user_role in ['Esperto', 'Admin']:
            if st.button("Consulta dati", use_container_width=True):
                st.info("Funzionalità 'Consulta dati' in sviluppo...")

        # Button 4: Gestione delle utenze (Visible to 'Admin')
        if st.session_state.user_role == 'Admin':
            if st.button("Gestione delle utenze", use_container_width=True):
                st.info("Funzionalità 'Gestione delle utenze' in sviluppo...")

        st.write("")

        # Logout button
        if st.button("Logout", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.user_role = None
            st.session_state.page = 'welcome'
            st.rerun()

def quotation_calculation_page():
    """
    Displays the page to upload the Excel file and start the quotation process.
    """
    if not st.session_state.logged_in:
        st.warning("Per accedere a questa pagina, devi prima effettuare il login.")
        st.session_state.page = 'login'
        st.rerun()
        return
    
    # Initialize session state for the result message and retry button
    if 'result_message' not in st.session_state:
        st.session_state.result_message = ""
    if 'show_retry_button' not in st.session_state:
        st.session_state.show_retry_button = False

    col_empty, col_header, _ = st.columns([0.05, 3, 1])
    with col_header:
        st.markdown("# Pass Broker")

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("Calcola Nuovo Preventivo")
        st.write("---")
        
        uploaded_file = st.file_uploader("Carica il file clienti_assicurazioni.xlsx", type=["xlsx"])
              
#        if uploaded_file:
#            st.success("File caricato con successo!")
#            if st.button("Calcola", use_container_width=True):
#                st.write("Avvio del processo di calcolo dei preventivi...")
#                
#                with st.spinner('Calcolo in corso... potrebbe richiedere qualche minuto.'):
#                    try:
#                        df_uploaded = pd.read_excel(uploaded_file)
#                        st.session_state.result_message = run_quotation_process(df_uploaded)
#                        st.success("Calcolo completato.")
#                        st.text(st.session_state.result_message )
#                    except Exception as e:
#                        st.error(f"Si è verificato un errore durante l'elaborazione: {e}")
        
        if uploaded_file:
            st.success("File caricato con successo!")
            if not st.session_state.show_retry_button:
                if st.button("Avvia Elaborazione", type="primary"):
                    st.session_state.show_retry_button = False
                    with st.spinner('Calcolo in corso... potrebbe richiedere qualche minuto.'):
                        try:
                            if uploaded_file is not None:
                                df_uploaded = pd.read_excel(uploaded_file, dtype={'Auto': str})
                                st.session_state.result_message = run_quotation_process(df_uploaded)
                                st.success("Calcolo completato.")
                                st.text(st.session_state.result_message)

                        except Exception as e:
                            st.session_state.result_message = f"Si è verificato un errore inaspettato durante il caricamento o l'elaborazione: {e}"
                    
                    # Check the result message for the presence of "KO" or "Errore"
                    if "KO" in st.session_state.result_message or "Errore" in st.session_state.result_message:
                        st.session_state.show_retry_button = True

            if st.session_state.show_retry_button:
                if st.button("Ricalcola", type="secondary"):
                    with st.spinner("Ricalcolo in corso, attendere prego..."):
                        try:
                            # Pass None to the function to default to the existing excel file
                            st.session_state.result_message = run_quotation_process()
                            st.success("Calcolo completato.")
                            st.text(st.session_state.result_message)
                        except Exception as e:
                            st.session_state.result_message = f"Errore inaspettato durante il ricalcolo: {e}"

        st.write("")
        if st.button("Torna alla Dashboard", use_container_width=True):
            st.session_state.page = 'dashboard'
            st.rerun()

def main():
    """
    Main function to manage the app's pages.
    """
    # Display the correct page based on the session state
    if st.session_state.page == 'welcome':
        welcome_page()
    elif st.session_state.page == 'login':
        login_form_page()
    elif st.session_state.page == 'dashboard':
        dashboard_page()
    elif st.session_state.page == 'quotation_calculation':
        quotation_calculation_page()

if __name__ == "__main__":
    main()
