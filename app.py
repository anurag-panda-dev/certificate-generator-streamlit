import streamlit as st
import requests
import json
from datetime import datetime, date

# Configure page
st.set_page_config(
    page_title="Certificate Generator",
    page_icon="🏆",
    layout="centered"
)

# Google Apps Script Web App URL - Replace with your actual URL
GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyvL9iPWYPHkwSsEGq4Eb7l_QHdSSvH5cMFEMbaSBOYElcS1GiRKH9z6CsEjL4uSzg/exec"

# Certificate types mapping
CERTIFICATE_TYPES = {
    "Python for Beginners": "Python",
    "Web Development for Beginners": "WebDev", 
    "Tree Plantation": "TreePlantation",
    "Debate Competition": "Debate",
    "Yoga Camp": "Yoga",
    "Blood Donation Camp": "BloodDonation",
    "Arduino for Beginners": "Arduino"
}

BLOOD_GROUPS = ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"]

# Initialize session state
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'selected_certificate' not in st.session_state:
    st.session_state.selected_certificate = None

def reset_form():
    """Reset the form to step 1"""
    st.session_state.step = 1
    st.session_state.selected_certificate = None

def submit_to_google_sheets(data):
    """Submit form data to Google Sheets via Apps Script"""
    try:
        # Convert date to string for JSON serialization
        if isinstance(data.get('date'), date):
            data['date'] = data['date'].strftime('%Y-%m-%d')
        
        response = requests.post(
            GOOGLE_APPS_SCRIPT_URL,
            json=data,
            headers={'Content-Type': 'application/json'}
        )
        
        if response.status_code == 200:
            return True, "Certificate request submitted successfully!"
        else:
            return False, f"Error: {response.status_code} - {response.text}"
    
    except Exception as e:
        return False, f"Error submitting form: {str(e)}"

# Main app layout
st.title("🏆 Certificate Generator")
st.markdown("---")

# Step 1: Certificate Type Selection
if st.session_state.step == 1:
    st.header("Step 1: Select Certificate Type")
    st.markdown("Choose the certificate you want to generate:")
    
    # Create a grid layout for certificate selection
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🐍 Python for Beginners", use_container_width=True):
            st.session_state.selected_certificate = "Python for Beginners"
            st.session_state.step = 2
            st.rerun()
        
        if st.button("🌐 Web Development for Beginners", use_container_width=True):
            st.session_state.selected_certificate = "Web Development for Beginners"
            st.session_state.step = 2
            st.rerun()
        
        if st.button("🌳 Tree Plantation", use_container_width=True):
            st.session_state.selected_certificate = "Tree Plantation"
            st.session_state.step = 2
            st.rerun()
        
        if st.button("🗣️ Debate Competition", use_container_width=True):
            st.session_state.selected_certificate = "Debate Competition"
            st.session_state.step = 2
            st.rerun()
    
    with col2:
        if st.button("🧘 Yoga Camp", use_container_width=True):
            st.session_state.selected_certificate = "Yoga Camp"
            st.session_state.step = 2
            st.rerun()
        
        if st.button("🩸 Blood Donation Camp", use_container_width=True):
            st.session_state.selected_certificate = "Blood Donation Camp"
            st.session_state.step = 2
            st.rerun()
        
        if st.button("🤖 Arduino for Beginners", use_container_width=True):
            st.session_state.selected_certificate = "Arduino for Beginners"
            st.session_state.step = 2
            st.rerun()

# Step 2: Form Fill
elif st.session_state.step == 2:
    st.header("Step 2: Fill Certificate Details")
    
    # Show selected certificate type
    st.info(f"Selected Certificate: **{st.session_state.selected_certificate}**")
    
    # Back button
    if st.button("← Back to Certificate Selection"):
        reset_form()
        st.rerun()
    
    st.markdown("---")
    
    # Form
    with st.form("certificate_form"):
        st.subheader("Certificate Information")
        
        # Full Name
        full_name = st.text_input(
            "Full Name *",
            placeholder="Enter your full name as it should appear on the certificate",
            help="This name will appear on your certificate"
        )
        
        # Email Address
        email = st.text_input(
            "Email Address *",
            placeholder="your.email@example.com",
            help="We'll send your certificate to this email address"
        )
        
        # Date of Issue
        issue_date = st.date_input(
            "Date of Issue *",
            value=datetime.now().date(),
            help="The date when the certificate should be issued"
        )
        
        # Blood Group (conditional field)
        blood_group = None
        if st.session_state.selected_certificate == "Blood Donation Camp":
            blood_group = st.selectbox(
                "Blood Group *",
                options=BLOOD_GROUPS,
                help="Select your blood group"
            )
        
        # Submit button
        submitted = st.form_submit_button("🚀 Generate Certificate", use_container_width=True)
        
        if submitted:
            # Validate required fields
            if not full_name or not email:
                st.error("Please fill in all required fields.")
            elif st.session_state.selected_certificate == "Blood Donation Camp" and not blood_group:
                st.error("Please select your blood group.")
            else:
                # Prepare data for submission
                form_data = {
                    "type": st.session_state.selected_certificate,
                    "name": full_name,
                    "email": email,
                    "date": issue_date,
                    "blood_group": blood_group if blood_group else ""
                }
                
                # Submit to Google Sheets
                with st.spinner("Submitting your certificate request..."):
                    success, message = submit_to_google_sheets(form_data)
                
                if success:
                    st.success(message)
                    st.balloons()
                    
                    # Show success message and option to create another certificate
                    st.markdown("---")
                    st.markdown("### What's Next?")
                    st.info("Your certificate request has been submitted successfully! You will receive your certificate via email shortly.")
                    
                    if st.button("Create Another Certificate", use_container_width=True):
                        reset_form()
                        st.rerun()
                else:
                    st.error(message)
                    st.markdown("---")
                    st.markdown("### Having Issues?")
                    st.warning("If you continue to experience problems, please contact the administrator.")

# Sidebar with information
with st.sidebar:
    st.header("ℹ️ Information")
    st.markdown("""
    **Available Certificates:**
    - 🐍 Python for Beginners
    - 🌐 Web Development for Beginners  
    - 🌳 Tree Plantation
    - 🗣️ Debate Competition
    - 🧘 Yoga Camp
    - 🩸 Blood Donation Camp
    - 🤖 Arduino for Beginners
    """)
    
    st.markdown("---")
    st.markdown("**How it works:**")
    st.markdown("""
    1. Select your certificate type
    2. Fill in your details
    3. Submit the form
    4. Receive your certificate via email
    """)
    
    st.markdown("---")
    st.markdown("**Need Help?**")
    st.markdown("Contact support if you encounter any issues.")

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666; font-size: 0.8em;'>"
    "© 2025 Certificate Generator | Built with Streamlit by Anurag Panda"
    "</div>",
    unsafe_allow_html=True
)
