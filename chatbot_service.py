import os
import openai
from typing import Dict, Any, List
from datetime import datetime
import pandas as pd
from sharepoint_items import (
    fetch_sharepoint_list,
    items_to_dataframe,
    get_user_analytics_specific,
    get_all_contacts_from_onedrive,
    get_all_customers_from_onedrive,
    get_user_details_from_excell,
    get_user_tasks_details_from_excell,
    compute_user_analytics_with_last_date,
    get_partnership_data,
    get_access_token,
    list_org_users
)

CHAT_HISTORY_LIMIT = 20
MODEL_NAME = "gpt-4"
SYSTEM_PROMPT = """You are a helpful assistant for Hamdaz employees. You have access to the user's data and can help with various tasks.

Guidelines:
1. Only use the data provided in the context
2. Be concise and professional
3. If you don't know the answer, say so
4. Never make up information
5. Always maintain user privacy and data security
6. CRITICAL: You can ONLY access data for the user you're currently helping
7. Never reveal information about other users unless you're helping an admin
"""

ADMIN_SYSTEM_PROMPT = """You are a helpful administrative assistant for Hamdaz. You have full access to all organizational data.

Guidelines:
1. You have access to all user data and organizational information
2. Be concise and professional
3. Provide comprehensive analytics and insights
4. Never make up information
5. Maintain data security and confidentiality
"""

class ChatbotService:
    def __init__(self):
        self.openai_api_key = os.getenv("OPENAI_API_KEY")
        openai.api_key = self.openai_api_key
        
        # Configuration
        self.SITE_DOMAIN = os.getenv("SITE_DOMAIN", "hamdaz1.sharepoint.com")
        self.SITE_PATH = os.getenv("SITE_PATH", "/sites/ProposalTeam")
        self.LIST_NAME = os.getenv("LIST_NAME", "Proposals")
        self.EXCLUDED_USERS = ["Sebin", "Shamshad", "Jaymon", "Hisham Arackal", "Althaf", "Nidal", "Nayif Muhammed S", "Afthab"]
        self.ADMIN_EMAILS = ["jishad@hamdaz.com", "sebin@hamdaz.com"]
    
    def is_admin(self, user_email: str) -> bool:
        """Check if user is an admin"""
        return user_email.lower().strip() in [email.lower().strip() for email in self.ADMIN_EMAILS]
    
    def get_user_context(self, user_email: str, is_admin: bool = False) -> Dict[str, Any]:
        """
        Fetch user-specific data to include in the chat context.
        Admins get all data, regular users get only their own data.
        """
        context = {
            "user_email": user_email,
            "is_admin": is_admin,
            "timestamp": datetime.now().isoformat(),
            "data_sources": {}
        }
        
        try:
            # Fetch SharePoint tasks
            all_tasks = fetch_sharepoint_list(self.SITE_DOMAIN, self.SITE_PATH, self.LIST_NAME)
            df_all_tasks = items_to_dataframe(all_tasks)
            
            if is_admin:
                # Admin gets ALL data
                context["data_sources"]["all_tasks"] = all_tasks
                context["data_sources"]["all_tasks_summary"] = {
                    "total_tasks": len(df_all_tasks),
                    "total_users": df_all_tasks['AssignedTo'].nunique() if 'AssignedTo' in df_all_tasks.columns else 0
                }
                
                # Get all user analytics
                user_analytics = compute_user_analytics_with_last_date(
                    df_all_tasks, 
                    self.EXCLUDED_USERS
                )
                context["data_sources"]["all_user_analytics"] = user_analytics
                
                # Get all contacts
                all_contacts = get_all_contacts_from_onedrive()
                context["data_sources"]["all_contacts"] = all_contacts
                context["data_sources"]["contacts_count"] = len(all_contacts)
                
                # Get all customers
                all_customers = get_all_customers_from_onedrive()
                context["data_sources"]["all_customers"] = all_customers
                context["data_sources"]["customers_count"] = len(all_customers)
                
                # Get all user details
                all_user_details = get_user_details_from_excell()
                context["data_sources"]["all_user_details"] = all_user_details
                
                # Get partnership data
                partnership_data = get_partnership_data()
                context["data_sources"]["partnership_data"] = partnership_data
                
                # Get org users list
                access_token = get_access_token()
                org_users = list_org_users(access_token)
                context["data_sources"]["org_users"] = org_users
                
            else:
                # Regular user gets ONLY their own data
                # Filter tasks for this specific user
                user_tasks = [task for task in all_tasks if task.get('AssignedTo') == user_email]
                context["data_sources"]["my_tasks"] = user_tasks
                context["data_sources"]["my_tasks_count"] = len(user_tasks)
                
                # Get user-specific analytics
                if not df_all_tasks.empty and 'AssignedTo' in df_all_tasks.columns:
                    user_analytics = get_user_analytics_specific(df_all_tasks, user_email)
                    context["data_sources"]["my_analytics"] = user_analytics
                
                # Get only contacts created/owned by this user
                all_contacts = get_all_contacts_from_onedrive()
                user_contacts = [c for c in all_contacts if c.get('CreatedBy') == user_email or c.get('Owner') == user_email]
                context["data_sources"]["my_contacts"] = user_contacts
                context["data_sources"]["my_contacts_count"] = len(user_contacts)
                
                # Get user's own profile details only
                all_user_details = get_user_details_from_excell()
                user_profile = next((u for u in all_user_details if u.get('email') == user_email), None)
                if user_profile:
                    context["data_sources"]["my_profile"] = user_profile
                
                # Partnership data - only if user has relevant assignments
                partnership_data = get_partnership_data()
                # Filter partnership data based on user assignments (if applicable)
                context["data_sources"]["partnership_data"] = partnership_data
            
            return context
            
        except Exception as e:
            print(f"Error fetching user context: {e}")
            context["error"] = str(e)
            return context
    
    def build_context_message(self, context: Dict[str, Any]) -> str:
        """
        Build a comprehensive context message for the LLM
        """
        is_admin = context.get("is_admin", False)
        user_email = context.get("user_email")
        
        if is_admin:
            context_msg = f"""You are helping an ADMINISTRATOR: {user_email}

ADMIN ACCESS - Full organizational data available:

"""
            data = context.get("data_sources", {})
            
            # Summary statistics
            context_msg += f"Total Tasks: {data.get('all_tasks_summary', {}).get('total_tasks', 0)}\n"
            context_msg += f"Total Active Users: {data.get('all_tasks_summary', {}).get('total_users', 0)}\n"
            context_msg += f"Total Contacts: {data.get('contacts_count', 0)}\n"
            context_msg += f"Total Customers: {data.get('customers_count', 0)}\n\n"
            
            # User analytics summary
            if data.get('all_user_analytics'):
                context_msg += "User Analytics Summary:\n"
                for user, analytics in list(data['all_user_analytics'].items())[:10]:
                    context_msg += f"  - {user}: {analytics.get('total_tasks', 0)} tasks, {analytics.get('tasks_pending', 0)} pending\n"
                context_msg += "\n"
            
            context_msg += "You have access to all organizational data. Provide comprehensive insights and analytics.\n"
            
        else:
            context_msg = f"""You are helping USER: {user_email}

IMPORTANT: You can ONLY access and discuss THIS USER'S data. Never reveal information about other users.

User's Personal Data:

"""
            data = context.get("data_sources", {})
            
            # User's tasks
            my_tasks = data.get('my_tasks', [])
            context_msg += f"My Tasks ({len(my_tasks)} total):\n"
            if my_tasks:
                for task in my_tasks[:5]:  # Show first 5
                    context_msg += f"  - {task.get('Title', 'N/A')} (Status: {task.get('SubmissionStatus', 'N/A')})\n"
                if len(my_tasks) > 5:
                    context_msg += f"  ... and {len(my_tasks) - 5} more tasks\n"
            context_msg += "\n"
            
            # User's analytics
            my_analytics = data.get('my_analytics', {})
            if my_analytics:
                context_msg += "My Performance:\n"
                context_msg += f"  - Total Tasks: {my_analytics.get('TotalTasks', 0)}\n"
                context_msg += f"  - Completed: {my_analytics.get('CompletedTasksCount', 0)}\n"
                context_msg += f"  - Ongoing: {my_analytics.get('OngoingTasksCount', 0)}\n"
                context_msg += f"  - Missed: {my_analytics.get('MissedTasksCount', 0)}\n\n"
            
            # User's contacts
            my_contacts_count = data.get('my_contacts_count', 0)
            context_msg += f"My Contacts: {my_contacts_count}\n\n"
            
            # User's profile
            my_profile = data.get('my_profile')
            if my_profile:
                context_msg += "My Profile:\n"
                context_msg += f"  - Name: {my_profile.get('name', 'N/A')}\n"
                context_msg += f"  - Role: {my_profile.get('role', 'N/A')}\n\n"
        
        context_msg += f"\nCurrent Time: {context.get('timestamp')}\n"
        return context_msg
    
    def format_chat_history(self, messages: List[Dict[str, str]]) -> List[Dict[str, str]]:
        """Format and limit chat history"""
        # Keep only the last N messages
        return messages[-CHAT_HISTORY_LIMIT:]
    
    def chat(self, user_email: str, user_message: str, chat_history: List[Dict[str, str]] = None) -> Dict[str, Any]:
        """
        Main chat function with strict user isolation
        """
        try:
            # Determine if user is admin
            is_admin = self.is_admin(user_email)
            
            # Get user-specific context
            context = self.get_user_context(user_email, is_admin)
            
            if "error" in context:
                return {
                    "success": False,
                    "error": f"Failed to fetch user context: {context['error']}",
                    "response": "I'm having trouble accessing your data. Please try again later."
                }
            
            # Build context message
            context_message = self.build_context_message(context)
            
            # Prepare messages for OpenAI
            messages = [
                {
                    "role": "system",
                    "content": ADMIN_SYSTEM_PROMPT if is_admin else SYSTEM_PROMPT
                },
                {
                    "role": "system",
                    "content": context_message
                }
            ]
            
            # Add chat history if provided
            if chat_history:
                formatted_history = self.format_chat_history(chat_history)
                messages.extend(formatted_history)
            
            # Add current user message
            messages.append({
                "role": "user",
                "content": user_message
            })
            
            # Call OpenAI API
            response = openai.ChatCompletion.create(
                model=MODEL_NAME,
                messages=messages,
                temperature=0.7,
                max_tokens=1000
            )
            
            assistant_message = response.choices[0].message.content
            
            return {
                "success": True,
                "response": assistant_message,
                "context_used": {
                    "is_admin": is_admin,
                    "data_sources": list(context.get("data_sources", {}).keys())
                },
                "tokens_used": response.usage.total_tokens
            }
            
        except Exception as e:
            print(f"Error in chat: {e}")
            return {
                "success": False,
                "error": str(e),
                "response": "I encountered an error processing your request. Please try again."
            }
    

# Example usage
if __name__ == "__main__":
    chatbot = ChatbotService()
    
    # Example: Regular user
    regular_user = "john.doe@hamdaz.com"
    response = chatbot.chat(
        user_email=regular_user,
        user_message="What are my pending tasks?"
    )
    print(f"Regular User Response: {response['response']}\n")
    
    # Example: Admin user
    admin_user = "admin@hamdaz.com"  # Make sure this is in ADMIN_EMAILS env var
    admin_response = chatbot.chat(
        user_email=admin_user,
        user_message="Give me a summary of team performance"
    )
    print(f"Admin Response: {admin_response['response']}")