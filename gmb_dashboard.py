import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import requests
import json
from datetime import datetime, timedelta
import io
import base64
from textblob import TextBlob
import google.auth
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import tempfile
import pickle
import time
from google.oauth2.credentials import Credentials

# Set page config
st.set_page_config(
    page_title="Google My Business Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main > div {
        padding-top: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        margin: 0.5rem 0;
    }
    .negative-sentiment {
        background-color: #ffebee;
        border-left: 4px solid #f44336;
        padding: 10px;
        margin: 5px 0;
    }
    .neutral-sentiment {
        background-color: #fff3e0;
        border-left: 4px solid #ff9800;
        padding: 10px;
        margin: 5px 0;
    }
    .positive-sentiment {
        background-color: #e8f5e8;
        border-left: 4px solid #4caf50;
        padding: 10px;
        margin: 5px 0;
    }
    .export-section {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

class GMBAnalytics:
    def __init__(self):
        self.service = None
        self.credentials = None
        self.SCOPES = ['https://www.googleapis.com/auth/business.manage']
        
    def get_oauth_flow(self):
        """Create OAuth2 flow"""
        try:
            # Get OAuth2 credentials from secrets
            oauth_config = {
                "web": {
                    "client_id": st.secrets["oauth2"]["client_id"],
                    "client_secret": st.secrets["oauth2"]["client_secret"],
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "redirect_uris": [st.secrets["oauth2"]["redirect_uri"]]
                }
            }
            
            flow = Flow.from_client_config(
                oauth_config,
                scopes=self.SCOPES,
                redirect_uri=st.secrets["oauth2"]["redirect_uri"]
            )
            return flow
        except Exception as e:
            st.error(f"Failed to create OAuth flow: {str(e)}")
            return None
    
    def get_auth_url(self):
        """Get authorization URL"""
        flow = self.get_oauth_flow()
        if flow:
            auth_url, _ = flow.authorization_url(
                access_type='offline',
                include_granted_scopes='true',
                prompt='consent'
            )
            return auth_url
        return None
    
    def authenticate_with_code(self, auth_code):
        """Authenticate using authorization code"""
        try:
            flow = self.get_oauth_flow()
            if not flow:
                return False
            
            # Exchange authorization code for credentials
            flow.fetch_token(code=auth_code)
            self.credentials = flow.credentials
            
            # Build the service
            self.service = build('mybusinessbusinessinformation', 'v1', 
                               credentials=self.credentials)
            
            # Store credentials in session state for persistence
            st.session_state['gmb_credentials'] = {
                'token': self.credentials.token,
                'refresh_token': self.credentials.refresh_token,
                'token_uri': self.credentials.token_uri,
                'client_id': self.credentials.client_id,
                'client_secret': self.credentials.client_secret,
                'scopes': self.credentials.scopes
            }
            
            return True
            
        except Exception as e:
            st.error(f"Authentication failed: {str(e)}")
            return False
    
    def load_credentials_from_session(self):
        """Load credentials from session state"""
        try:
            if 'gmb_credentials' in st.session_state:
                cred_data = st.session_state['gmb_credentials']
                self.credentials = Credentials(
                    token=cred_data['token'],
                    refresh_token=cred_data['refresh_token'],
                    token_uri=cred_data['token_uri'],
                    client_id=cred_data['client_id'],
                    client_secret=cred_data['client_secret'],
                    scopes=cred_data['scopes']
                )
                
                # Refresh if needed
                if self.credentials.expired:
                    self.credentials.refresh(Request())
                    # Update session state with new token
                    st.session_state['gmb_credentials']['token'] = self.credentials.token
                
                # Build services
                self.service = build('mybusinessbusinessinformation', 'v1', 
                                   credentials=self.credentials)
                
                return True
        except Exception as e:
            st.error(f"Failed to load credentials: {str(e)}")
            if 'gmb_credentials' in st.session_state:
                del st.session_state['gmb_credentials']
        
        return False
    
    def get_accounts(self):
        """Get Google My Business accounts"""
        try:
            if not self.service:
                return []
            
            # Call the API to get accounts
            accounts = self.service.accounts().list().execute()
            return accounts.get('accounts', [])
            
        except Exception as e:
            st.error(f"Failed to get accounts: {str(e)}")
            return []
    
    def get_locations(self):
        """Get all business locations"""
        try:
            if not self.service:
                return []
            
            all_locations = []
            accounts = self.get_accounts()
            
            for account in accounts:
                account_name = account['name']
                
                # Get locations for this account
                try:
                    locations_response = self.service.accounts().locations().list(
                        parent=account_name
                    ).execute()
                    
                    locations = locations_response.get('locations', [])
                    
                    for location in locations:
                        all_locations.append({
                            'id': location['name'],
                            'name': location.get('title', 'Unknown Location'),
                            'brand': account.get('accountName', 'Unknown Brand'),
                            'address': self._format_address(location.get('storefrontAddress', {})),
                            'phone': location.get('primaryPhone', ''),
                            'website': location.get('websiteUri', ''),
                            'account_id': account_name
                        })
                        
                except Exception as loc_error:
                    st.warning(f"Could not fetch locations for account {account.get('accountName', 'Unknown')}: {str(loc_error)}")
                    continue
            
            return all_locations
            
        except Exception as e:
            st.error(f"Failed to get locations: {str(e)}")
            # Return demo data as fallback
            return self._get_demo_locations()
    
    def _format_address(self, address):
        """Format address object to string"""
        if not address:
            return ""
        
        parts = []
        if address.get('addressLines'):
            parts.extend(address['addressLines'])
        if address.get('locality'):
            parts.append(address['locality'])
        if address.get('administrativeArea'):
            parts.append(address['administrativeArea'])
        if address.get('postalCode'):
            parts.append(address['postalCode'])
            
        return ', '.join(parts)
    
    def get_insights(self, location_id, start_date, end_date):
        """Get insights data for a location"""
        try:
            if not self.service:
                return self._get_demo_insights(location_id, start_date, end_date)
            
            # Build insights service
            insights_service = build('mybusinessbusinessinformation', 'v1', 
                                   credentials=self.credentials)
            
            # Format dates for API
            start_date_str = start_date.strftime('%Y-%m-%d')
            end_date_str = end_date.strftime('%Y-%m-%d')
            
            # Request insights data
            request_body = {
                'locationNames': [location_id],
                'basicRequest': {
                    'metricRequests': [
                        {'metric': 'QUERIES_DIRECT'},
                        {'metric': 'QUERIES_INDIRECT'},
                        {'metric': 'VIEWS_MAPS'},
                        {'metric': 'VIEWS_SEARCH'},
                        {'metric': 'ACTIONS_WEBSITE'},
                        {'metric': 'ACTIONS_PHONE'},
                        {'metric': 'ACTIONS_DRIVING_DIRECTIONS'},
                        {'metric': 'PHOTOS_VIEWS_MERCHANT'},
                        {'metric': 'PHOTOS_VIEWS_CUSTOMERS'}
                    ],
                    'timeRange': {
                        'startTime': start_date_str + 'T00:00:00Z',
                        'endTime': end_date_str + 'T23:59:59Z'
                    }
                }
            }
            
            # Make API call to get insights
            insights_response = insights_service.accounts().locations().getInsights(
                name=location_id,
                body=request_body
            ).execute()
            
            # Process the response into DataFrame
            return self._process_insights_response(insights_response, location_id, start_date, end_date)
            
        except Exception as e:
            st.warning(f"Could not fetch real insights data: {str(e)}. Using demo data.")
            return self._get_demo_insights(location_id, start_date, end_date)
    
    def _process_insights_response(self, response, location_id, start_date, end_date):
        """Process insights API response into DataFrame"""
        try:
            # Extract metrics from response
            location_metrics = response.get('locationMetrics', [])
            if not location_metrics:
                return self._get_demo_insights(location_id, start_date, end_date)
            
            metrics = location_metrics[0].get('metricValues', [])
            
            # Create daily data structure
            days = (end_date - start_date).days + 1
            dates = [start_date + timedelta(days=i) for i in range(days)]
            
            # Initialize data dictionary
            data = {
                'date': dates,
                'search_impressions': [0] * days,
                'map_impressions': [0] * days,
                'website_clicks': [0] * days,
                'direction_requests': [0] * days,
                'phone_calls': [0] * days,
                'photo_views': [0] * days,
                'location_id': [location_id] * days
            }
            
            # Process each metric
            for metric in metrics:
                metric_name = metric.get('metric', '')
                daily_values = metric.get('dimensionalValues', [])
                
                for i, daily_value in enumerate(daily_values[:days]):
                    value = int(daily_value.get('value', 0))
                    
                    if metric_name in ['VIEWS_SEARCH', 'QUERIES_DIRECT', 'QUERIES_INDIRECT']:
                        data['search_impressions'][i] += value
                    elif metric_name == 'VIEWS_MAPS':
                        data['map_impressions'][i] = value
                    elif metric_name == 'ACTIONS_WEBSITE':
                        data['website_clicks'][i] = value
                    elif metric_name == 'ACTIONS_DRIVING_DIRECTIONS':
                        data['direction_requests'][i] = value
                    elif metric_name == 'ACTIONS_PHONE':
                        data['phone_calls'][i] = value
                    elif metric_name in ['PHOTOS_VIEWS_MERCHANT', 'PHOTOS_VIEWS_CUSTOMERS']:
                        data['photo_views'][i] += value
            
            return pd.DataFrame(data)
            
        except Exception as e:
            st.warning(f"Error processing insights response: {str(e)}")
            return self._get_demo_insights(location_id, start_date, end_date)
    
    def get_reviews(self, location_id):
        """Get reviews for a location"""
        try:
            if not self.service:
                return self._get_demo_reviews(location_id)
            
            # Build reviews service 
            reviews_service = build('mybusinessbusinessinformation', 'v1', 
                                  credentials=self.credentials)
            
            # Get reviews
            reviews_response = reviews_service.accounts().locations().reviews().list(
                parent=location_id
            ).execute()
            
            reviews_data = []
            reviews = reviews_response.get('reviews', [])
            
            for i, review in enumerate(reviews):
                reviews_data.append({
                    'id': i + 1,
                    'author': review.get('reviewer', {}).get('displayName', 'Anonymous'),
                    'rating': review.get('starRating', 0),
                    'text': review.get('comment', ''),
                    'date': review.get('createTime', '')[:10] if review.get('createTime') else '',
                    'reply': review.get('reviewReply', {}).get('comment', '') if review.get('reviewReply') else ''
                })
            
            return pd.DataFrame(reviews_data)
            
        except Exception as e:
            st.warning(f"Could not fetch real reviews: {str(e)}. Using demo data.")
            return self._get_demo_reviews(location_id)
    
    def _get_demo_locations(self):
        """Return demo locations data"""
        return [
            {"id": "loc_1", "name": "Downtown Store", "brand": "Brand A", "address": "123 Main St, City, ST 12345", "phone": "(555) 123-4567"},
            {"id": "loc_2", "name": "Mall Location", "brand": "Brand A", "address": "456 Mall Ave, City, ST 12345", "phone": "(555) 234-5678"},
            {"id": "loc_3", "name": "Airport Store", "brand": "Brand B", "address": "789 Airport Rd, City, ST 12345", "phone": "(555) 345-6789"},
            {"id": "loc_4", "name": "Suburb Branch", "brand": "Brand B", "address": "321 Suburb Ln, City, ST 12345", "phone": "(555) 456-7890"},
        ]
    
    def _get_demo_insights(self, location_id, start_date, end_date):
        """Return demo insights data"""
        import random
        import numpy as np
        
        days = (end_date - start_date).days + 1
        dates = [start_date + timedelta(days=i) for i in range(days)]
        
        return pd.DataFrame({
            'date': dates,
            'search_impressions': np.random.randint(50, 200, days),
            'map_impressions': np.random.randint(30, 150, days),
            'website_clicks': np.random.randint(5, 50, days),
            'direction_requests': np.random.randint(10, 80, days),
            'phone_calls': np.random.randint(2, 25, days),
            'photo_views': np.random.randint(20, 100, days),
            'location_id': location_id
        })
    
    def _get_demo_reviews(self, location_id):
        """Return demo reviews data"""
        reviews = [
            {"id": 1, "author": "John D.", "rating": 5, "text": "Excellent service! The staff was very helpful and friendly. Will definitely come back.", "date": "2024-07-20"},
            {"id": 2, "author": "Sarah M.", "rating": 4, "text": "Good experience overall. The product quality is great but the wait time was a bit long.", "date": "2024-07-18"},
            {"id": 3, "author": "Mike R.", "rating": 2, "text": "Disappointing visit. The staff seemed uninterested and the place was messy.", "date": "2024-07-15"},
            {"id": 4, "author": "Lisa K.", "rating": 5, "text": "Amazing! Best customer service I've experienced. Highly recommend this place.", "date": "2024-07-12"},
            {"id": 5, "author": "Tom W.", "rating": 3, "text": "It's okay, nothing special. Average service and products.", "date": "2024-07-10"},
            {"id": 6, "author": "Emma B.", "rating": 1, "text": "Terrible experience. Rude staff and poor quality products. Won't be returning.", "date": "2024-07-08"},
        ]
        return pd.DataFrame(reviews)

def analyze_sentiment(text):
    """Analyze sentiment of review text"""
    blob = TextBlob(text)
    polarity = blob.sentiment.polarity
    
    if polarity > 0.1:
        return 'Positive'
    elif polarity < -0.1:
        return 'Negative'
    else:
        return 'Neutral'

def generate_sentiment_summary(reviews_df):
    """Generate summaries for each sentiment category"""
    summaries = {}
    
    for sentiment in ['Positive', 'Negative', 'Neutral']:
        sentiment_reviews = reviews_df[reviews_df['sentiment'] == sentiment]
        if len(sentiment_reviews) > 0:
            # Extract common themes/keywords
            all_text = ' '.join(sentiment_reviews['text'].tolist())
            blob = TextBlob(all_text)
            
            # Simple keyword extraction (in production, use more sophisticated NLP)
            words = [word.lower() for word in blob.words if len(word) > 3]
            word_freq = pd.Series(words).value_counts().head(5)
            
            summaries[sentiment] = {
                'count': len(sentiment_reviews),
                'avg_rating': sentiment_reviews['rating'].mean(),
                'keywords': word_freq.index.tolist(),
                'sample_review': sentiment_reviews.iloc[0]['text'] if len(sentiment_reviews) > 0 else ""
            }
    
    return summaries

def create_pdf_report(data, filename):
    """Create PDF report"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    
    # Title
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#2E86AB'),
        alignment=1  # Center alignment
    )
    story.append(Paragraph("Google My Business Analytics Report", title_style))
    story.append(Spacer(1, 12))
    
    # Date range
    story.append(Paragraph(f"Report Period: {data.get('date_range', 'N/A')}", styles['Normal']))
    story.append(Spacer(1, 12))
    
    # Summary metrics
    if 'summary_metrics' in data:
        story.append(Paragraph("Summary Metrics", styles['Heading2']))
        metrics_data = [['Metric', 'Value']]
        for key, value in data['summary_metrics'].items():
            metrics_data.append([key.replace('_', ' ').title(), str(value)])
        
        metrics_table = Table(metrics_data)
        metrics_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E86AB')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(metrics_table)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

def export_to_excel(data, filename):
    """Export data to Excel"""
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Summary sheet
        if 'summary_metrics' in data:
            summary_df = pd.DataFrame(list(data['summary_metrics'].items()), 
                                    columns=['Metric', 'Value'])
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Insights data
        if 'insights_data' in data:
            data['insights_data'].to_excel(writer, sheet_name='Insights', index=False)
        
        # Reviews data
        if 'reviews_data' in data:
            data['reviews_data'].to_excel(writer, sheet_name='Reviews', index=False)
    
    buffer.seek(0)
    return buffer

def main():
    # Initialize session state
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'gmb_analytics' not in st.session_state:
        st.session_state.gmb_analytics = GMBAnalytics()
    
    # Header
    st.title("📊 Google My Business Analytics Dashboard")
    st.markdown("---")
    
    # Authentication Section
    if not st.session_state.authenticated:
        st.header("🔐 Google My Business Authentication")
        
        # Check if we can load existing credentials
        gmb = st.session_state.gmb_analytics
        if gmb.load_credentials_from_session():
            st.session_state.authenticated = True
            st.rerun()
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.markdown("""
            ### Connect Your Google My Business Account
            
            To use this dashboard with real data, authenticate with your Google My Business account.
            
            **Setup Steps:**
            1. Go to [Google Cloud Console](https://console.cloud.google.com/)
            2. Create/select project and enable Google My Business APIs
            3. Create OAuth2 credentials for web application
            4. Add your Streamlit app URL as authorized redirect URI
            5. Configure secrets in Streamlit Cloud
            """)
            
            # OAuth2 Authentication Flow
            st.subheader("🔗 OAuth2 Authentication")
            
            # Check if we have OAuth2 configured
            try:
                oauth_configured = (
                    "oauth2" in st.secrets and 
                    "client_id" in st.secrets["oauth2"] and 
                    "client_secret" in st.secrets["oauth2"]
                )
            except:
                oauth_configured = False
            
            if oauth_configured:
                # Step 1: Get authorization URL
                if st.button("🚀 Start Google Authentication", type="primary"):
                    auth_url = gmb.get_auth_url()
                    if auth_url:
                        st.markdown(f"""
                        **Step 1:** Click the link below to authorize this app:
                        
                        [🔐 Authorize with Google My Business]({auth_url})
                        
                        **Step 2:** After authorization, copy the code from the redirect URL and paste it below.
                        """)
                        st.session_state.show_auth_code_input = True
                
                # Step 2: Handle authorization code
                if st.session_state.get('show_auth_code_input', False):
                    st.markdown("---")
                    auth_code = st.text_input(
                        "📋 Paste the authorization code here:",
                        placeholder="Enter the code from the redirect URL...",
                        help="After clicking the authorization link above, you'll be redirected to a URL with a 'code' parameter. Copy and paste that code here."
                    )
                    
                    if auth_code and st.button("✅ Complete Authentication"):
                        with st.spinner("Authenticating with Google My Business..."):
                            if gmb.authenticate_with_code(auth_code):
                                st.success("✅ Successfully authenticated with Google My Business!")
                                st.session_state.authenticated = True
                                st.session_state.show_auth_code_input = False
                                st.balloons()
                                time.sleep(2)
                                st.rerun()
                            else:
                                st.error("❌ Authentication failed. Please check your authorization code and try again.")
            
            else:
                st.error("""
                ❌ **OAuth2 not configured**
                
                Please configure your OAuth2 credentials in Streamlit secrets:
                
                ```toml
                [oauth2]
                client_id = "your-client-id.apps.googleusercontent.com"
                client_secret = "your-client-secret"
                redirect_uri = "https://your-app.streamlit.app"
                ```
                """)
        
        with col2:
            st.info("""
            **Demo Mode Available**
            
            Explore the dashboard with sample data without authentication.
            """)
            
            if st.button("📊 Use Demo Data", type="secondary"):
                st.session_state.authenticated = True
                st.session_state.demo_mode = True
                st.success("✅ Demo mode activated!")
                st.rerun()
            
            st.markdown("---")
            
            st.markdown("""
            **🔧 Need Help?**
            
            - [Google My Business API Setup Guide](https://developers.google.com/my-business/content/prereqs)
            - [OAuth2 Setup Instructions](https://developers.google.com/identity/protocols/oauth2)
            - Check the README for detailed setup steps
            """)
        
        return
    
    # Main Dashboard
    if st.session_state.get('demo_mode', False):
        st.info("📊 **Demo Mode Active** - Using sample data for demonstration")
    else:
        st.success("✅ Connected to Google My Business API")
        
        # Display connected account info
        try:
            accounts = st.session_state.gmb_analytics.get_accounts()
            if accounts:
                account_names = [acc.get('accountName', 'Unknown') for acc in accounts[:3]]
                st.caption(f"Connected accounts: {', '.join(account_names)}")
        except:
            pass
    
    # Sidebar Controls
    st.sidebar.header("📋 Dashboard Controls")
    
    # Get locations
    locations = st.session_state.gmb_analytics.get_locations()
    location_df = pd.DataFrame(locations)
    
    # Brand filter
    brands = ['All Brands'] + sorted(location_df['brand'].unique().tolist())
    selected_brand = st.sidebar.selectbox("🏷️ Select Brand", brands)
    
    # Filter locations by brand
    if selected_brand != 'All Brands':
        filtered_locations = location_df[location_df['brand'] == selected_brand]
    else:
        filtered_locations = location_df
    
    # Location filter
    location_options = ['All Locations'] + filtered_locations['name'].tolist()
    selected_location = st.sidebar.selectbox("📍 Select Location", location_options)
    
    # Date range
    st.sidebar.subheader("📅 Date Range")
    col1, col2 = st.sidebar.columns(2)
    with col1:
        start_date = st.date_input("Start Date", value=datetime.now() - timedelta(days=30))
    with col2:
        end_date = st.date_input("End Date", value=datetime.now())
    
    # Export section
    st.sidebar.markdown("---")
    st.sidebar.subheader("📤 Export Data")
    
    # Initialize data variables
    total_insights = pd.DataFrame()
    all_reviews = pd.DataFrame()
    
    # Main content area
    if selected_location == 'All Locations':
        # Summary view for all locations
        st.header("📈 Overview Dashboard")
        
        # Aggregate metrics
        for _, location in filtered_locations.iterrows():
            insights = st.session_state.gmb_analytics.get_insights(
                location['id'], start_date, end_date
            )
            insights['location_name'] = location['name']
            insights['brand'] = location['brand']
            total_insights = pd.concat([total_insights, insights], ignore_index=True)
            
            reviews = st.session_state.gmb_analytics.get_reviews(location['id'])
            reviews['location_name'] = location['name']
            reviews['brand'] = location['brand']
            all_reviews = pd.concat([all_reviews, reviews], ignore_index=True)
        
        # Summary metrics
        col1, col2, col3, col4 = st.columns(4)
        
        if not total_insights.empty:
            with col1:
                total_search = total_insights['search_impressions'].sum()
                st.metric("🔍 Total Search Impressions", f"{total_search:,}")
            
            with col2:
                total_map = total_insights['map_impressions'].sum()
                st.metric("🗺️ Total Map Impressions", f"{total_map:,}")
            
            with col3:
                total_clicks = total_insights['website_clicks'].sum()
                st.metric("👆 Website Clicks", f"{total_clicks:,}")
            
            with col4:
                total_calls = total_insights['phone_calls'].sum()
                st.metric("📞 Phone Calls", f"{total_calls:,}")
        
        # Charts
        if not total_insights.empty:
            st.subheader("📊 Performance Trends")
            
            # Daily trends
            daily_summary = total_insights.groupby('date').agg({
                'search_impressions': 'sum',
                'map_impressions': 'sum',
                'website_clicks': 'sum',
                'direction_requests': 'sum',
                'phone_calls': 'sum'
            }).reset_index()
            
            fig = make_subplots(
                rows=2, cols=2,
                subplot_titles=('Search vs Map Impressions', 'Customer Actions', 
                              'Performance by Location', 'Brand Comparison'),
                specs=[[{"secondary_y": False}, {"secondary_y": False}],
                       [{"secondary_y": False}, {"secondary_y": False}]]
            )
            
            # Impressions chart
            fig.add_trace(
                go.Scatter(x=daily_summary['date'], y=daily_summary['search_impressions'],
                          name='Search Impressions', line=dict(color='#1f77b4')),
                row=1, col=1
            )
            fig.add_trace(
                go.Scatter(x=daily_summary['date'], y=daily_summary['map_impressions'],
                          name='Map Impressions', line=dict(color='#ff7f0e')),
                row=1, col=1
            )
            
            # Actions chart
            fig.add_trace(
                go.Bar(x=['Website Clicks', 'Directions', 'Phone Calls'],
                      y=[daily_summary['website_clicks'].sum(),
                         daily_summary['direction_requests'].sum(),
                         daily_summary['phone_calls'].sum()],
                      name='Actions', marker_color=['#2ca02c', '#d62728', '#9467bd']),
                row=1, col=2
            )
            
            # Performance by location
            location_summary = total_insights.groupby('location_name')['search_impressions'].sum().reset_index()
            fig.add_trace(
                go.Bar(x=location_summary['location_name'], y=location_summary['search_impressions'],
                      name='Search Impressions by Location', marker_color='#17becf'),
                row=2, col=1
            )
            
            # Brand comparison
            brand_summary = total_insights.groupby('brand').agg({
                'search_impressions': 'sum',
                'map_impressions': 'sum'
            }).reset_index()
            fig.add_trace(
                go.Bar(x=brand_summary['brand'], y=brand_summary['search_impressions'],
                      name='Search by Brand', marker_color='#bcbd22'),
                row=2, col=2
            )
            
            fig.update_layout(height=800, showlegend=True)
            st.plotly_chart(fig, use_container_width=True)
        
        # Reviews analysis
        if not all_reviews.empty:
            st.subheader("⭐ Reviews Analysis")
            
            # Add sentiment analysis
            all_reviews['sentiment'] = all_reviews['text'].apply(analyze_sentiment)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                avg_rating = all_reviews['rating'].mean()
                st.metric("Average Rating", f"{avg_rating:.1f}⭐")
            
            with col2:
                total_reviews = len(all_reviews)
                st.metric("Total Reviews", total_reviews)
            
            with col3:
                recent_reviews = len(all_reviews[all_reviews['date'] >= str(start_date)])
                st.metric("Recent Reviews", recent_reviews)
            
            # Sentiment analysis
            sentiment_summary = generate_sentiment_summary(all_reviews)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if 'Positive' in sentiment_summary:
                    data = sentiment_summary['Positive']
                    st.markdown(f"""
                    <div class="positive-sentiment">
                        <h4>😊 Positive Reviews ({data['count']})</h4>
                        <p><strong>Avg Rating:</strong> {data['avg_rating']:.1f}⭐</p>
                        <p><strong>Key themes:</strong> {', '.join(data['keywords'][:3])}</p>
                    </div>
                    """, unsafe_allow_html=True)
            
            with col2:
                if 'Neutral' in sentiment_summary:
                    data = sentiment_summary['Neutral']
                    st.markdown(f"""
                    <div class="neutral-sentiment">
                        <h4>😐 Neutral Reviews ({data['count']})</h4>
                        <p><strong>Avg Rating:</strong> {data['avg_rating']:.1f}⭐</p>
                        <p><strong>Key themes:</strong> {', '.join(data['keywords'][:3])}</p>
                    </div>
                    """, unsafe_allow_html=True)
            
            with col3:
                if 'Negative' in sentiment_summary:
                    data = sentiment_summary['Negative']
                    st.markdown(f"""
                    <div class="negative-sentiment">
                        <h4>😞 Negative Reviews ({data['count']})</h4>
                        <p><strong>Avg Rating:</strong> {data['avg_rating']:.1f}⭐</p>
                        <p><strong>Key themes:</strong> {', '.join(data['keywords'][:3])}</p>
                    </div>
                    """, unsafe_allow_html=True)
            
            # Reviews distribution
            rating_dist = all_reviews['rating'].value_counts().sort_index()
            fig = px.bar(x=rating_dist.index, y=rating_dist.values,
                        title="Rating Distribution",
                        labels={'x': 'Rating', 'y': 'Count'},
                        color=rating_dist.values,
                        color_continuous_scale='RdYlGn')
            st.plotly_chart(fig, use_container_width=True)
    
    else:
        # Individual location view
        location_info = filtered_locations[filtered_locations['name'] == selected_location].iloc[0]
        st.header(f"📍 {selected_location} Analytics")
        st.caption(f"Brand: {location_info['brand']}")
        
        # Get data for selected location
        insights = st.session_state.gmb_analytics.get_insights(
            location_info['id'], start_date, end_date
        )
        reviews = st.session_state.gmb_analytics.get_reviews(location_info['id'])
        
        # Metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_search = insights['search_impressions'].sum()
            avg_search = insights['search_impressions'].mean()
            st.metric("🔍 Search Impressions", f"{total_search:,}", f"{avg_search:.1f}/day")
        
        with col2:
            total_map = insights['map_impressions'].sum()
            avg_map = insights['map_impressions'].mean()
            st.metric("🗺️ Map Impressions", f"{total_map:,}", f"{avg_map:.1f}/day")
        
        with col3:
            total_clicks = insights['website_clicks'].sum()
            st.metric("👆 Website Clicks", f"{total_clicks:,}")
        
        with col4:
            total_calls = insights['phone_calls'].sum()
            st.metric("📞 Phone Calls", f"{total_calls:,}")
        
        # Detailed charts
        st.subheader("📈 Performance Trends")
        
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=('Daily Impressions', 'Customer Actions Trend',
                          'Actions Distribution', 'Weekly Pattern'),
            specs=[[{"secondary_y": True}, {"secondary_y": False}],
                   [{"secondary_y": False}, {"secondary_y": False}]]
        )
        
        # Daily impressions with dual axis
        fig.add_trace(
            go.Scatter(x=insights['date'], y=insights['search_impressions'],
                      name='Search Impressions', line=dict(color='blue')),
            row=1, col=1, secondary_y=False
        )
        fig.add_trace(
            go.Scatter(x=insights['date'], y=insights['map_impressions'],
                      name='Map Impressions', line=dict(color='orange')),
            row=1, col=1, secondary_y=True
        )
        
        # Customer actions trend
        fig.add_trace(
            go.Scatter(x=insights['date'], y=insights['website_clicks'],
                      name='Website Clicks', line=dict(color='green')),
            row=1, col=2
        )
        fig.add_trace(
            go.Scatter(x=insights['date'], y=insights['phone_calls'],
                      name='Phone Calls', line=dict(color='red')),
            row=1, col=2
        )
        
        # Actions pie chart
        actions_data = {
            'Website Clicks': insights['website_clicks'].sum(),
            'Direction Requests': insights['direction_requests'].sum(),
            'Phone Calls': insights['phone_calls'].sum(),
            'Photo Views': insights['photo_views'].sum()
        }
        fig.add_trace(
            go.Pie(labels=list(actions_data.keys()), values=list(actions_data.values()),
                  name="Actions Distribution"),
            row=2, col=1
        )
        
        # Weekly pattern
        insights['weekday'] = pd.to_datetime(insights['date']).dt.day_name()
        weekly_avg = insights.groupby('weekday')['search_impressions'].mean().reindex([
            'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'
        ])
        fig.add_trace(
            go.Bar(x=weekly_avg.index, y=weekly_avg.values,
                  name='Avg Search Impressions by Day', marker_color='lightblue'),
            row=2, col=2
        )
        
        fig.update_layout(height=800, showlegend=True)
        st.plotly_chart(fig, use_container_width=True)
        
        # Reviews section
        st.subheader("⭐ Customer Reviews")
        
        if not reviews.empty:
            reviews['sentiment'] = reviews['text'].apply(analyze_sentiment)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                avg_rating = reviews['rating'].mean()
                st.metric("Average Rating", f"{avg_rating:.1f}⭐")
            
            with col2:
                total_reviews = len(reviews)
                st.metric("Total Reviews", total_reviews)
            
            with col3:
                latest_review_date = reviews['date'].max()
                st.metric("Latest Review", latest_review_date)
            
            # Sentiment analysis for this location
            sentiment_summary = generate_sentiment_summary(reviews)
            
            # Display sentiment summaries
            for sentiment, data in sentiment_summary.items():
                sentiment_class = f"{sentiment.lower()}-sentiment"
                emoji = {"Positive": "😊", "Neutral": "😐", "Negative": "😞"}[sentiment]
                
                st.markdown(f"""
                <div class="{sentiment_class}">
                    <h4>{emoji} {sentiment} Reviews ({data['count']})</h4>
                    <p><strong>Average Rating:</strong> {data['avg_rating']:.1f}⭐</p>
                    <p><strong>Common themes:</strong> {', '.join(data['keywords'])}</p>
                    <p><strong>Sample:</strong> "{data['sample_review'][:100]}..."</p>
                </div>
                """, unsafe_allow_html=True)
            
            # Recent reviews table
            st.subheader("📝 Recent Reviews")
            recent_reviews = reviews.sort_values('date', ascending=False).head(10)
            
            for _, review in recent_reviews.iterrows():
                sentiment_color = {
                    'Positive': '#4caf50',
                    'Neutral': '#ff9800', 
                    'Negative': '#f44336'
                }[review['sentiment']]
                
                st.markdown(f"""
                <div style="border-left: 4px solid {sentiment_color}; padding: 10px; margin: 10px 0; background-color: #f9f9f9;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <strong>{review['author']}</strong>
                        <span>{'⭐' * review['rating']} ({review['date']})</span>
                    </div>
                    <p style="margin: 5px 0 0 0;">{review['text']}</p>
                    <small style="color: {sentiment_color};">Sentiment: {review['sentiment']}</small>
                </div>
                """, unsafe_allow_html=True)
    
    # Export functionality in sidebar
    if st.sidebar.button("📊 Generate PDF Report"):
        # Prepare data for export
        export_data = {
            'date_range': f"{start_date} to {end_date}",
            'summary_metrics': {
                'total_search_impressions': total_insights['search_impressions'].sum() if not total_insights.empty else 0,
                'total_map_impressions': total_insights['map_impressions'].sum() if not total_insights.empty else 0,
                'total_website_clicks': total_insights['website_clicks'].sum() if not total_insights.empty else 0,
                'total_phone_calls': total_insights['phone_calls'].sum() if not total_insights.empty else 0,
                'average_rating': all_reviews['rating'].mean() if not all_reviews.empty else 0,
                'total_reviews': len(all_reviews) if not all_reviews.empty else 0
            }
        }
        
        pdf_buffer = create_pdf_report(export_data, "gmb_report.pdf")
        
        st.sidebar.download_button(
            label="📥 Download PDF Report",
            data=pdf_buffer,
            file_name=f"gmb_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mime="application/pdf"
        )
    
    if st.sidebar.button("📈 Generate Excel Report"):
        # Prepare data for Excel export
        export_data = {
            'insights_data': total_insights if not total_insights.empty else pd.DataFrame(),
            'reviews_data': all_reviews if not all_reviews.empty else pd.DataFrame(),
            'summary_metrics': {
                'total_search_impressions': total_insights['search_impressions'].sum() if not total_insights.empty else 0,
                'total_map_impressions': total_insights['map_impressions'].sum() if not total_insights.empty else 0,
                'total_website_clicks': total_insights['website_clicks'].sum() if not total_insights.empty else 0,
                'total_phone_calls': total_insights['phone_calls'].sum() if not total_insights.empty else 0,
                'average_rating': all_reviews['rating'].mean() if not all_reviews.empty else 0,
                'total_reviews': len(all_reviews) if not all_reviews.empty else 0
            }
        }
        
        excel_buffer = export_to_excel(export_data, "gmb_data.xlsx")
        
        st.sidebar.download_button(
            label="📥 Download Excel Report",
            data=excel_buffer,
            file_name=f"gmb_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p>📊 Google My Business Analytics Dashboard | Built with Streamlit</p>
        <p>For support or feature requests, contact your development team</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
