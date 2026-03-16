import os
import json
import openai
from duckduckgo_search import DDGS
import pandas as pd
from datetime import datetime
from cosmos import search_quotes_by_item
from sharepoint_items import fetch_sharepoint_list

SITE_DOMAIN = "hamdaz1.sharepoint.com"
test_path = "/sites/Test"
test_proposals_list = "testproposals"

def get_user_tasks(current_username, is_admin_user=False, target_username=None):
    """Fetches tasks from SharePoint."""
    try:
        tasks = fetch_sharepoint_list(SITE_DOMAIN, test_path, test_proposals_list)
        if is_admin_user:
            if target_username and target_username.lower() != "all":
                user_tasks = [t for t in tasks if str(t.get("AssignedTo", "")).replace(" ", "").lower() == target_username.lower()]
                if not user_tasks:
                    return f"No tasks found for user: {target_username}"
                return json.dumps(user_tasks, default=str)
            else:
                return json.dumps(tasks, default=str)
        else:
            user_tasks = [t for t in tasks if str(t.get("AssignedTo", "")).replace(" ", "") == current_username]
            if not user_tasks:
                return "No tasks found for the current user."
            return json.dumps(user_tasks, default=str)
    except Exception as e:
        return f"Error fetching tasks: {str(e)}"

def search_cosmos_db(query):
    """Searches local Cosmos DB for prices and product details."""
    try:
        results = search_quotes_by_item(query)
        if isinstance(results, pd.DataFrame) and results.empty:
            return "No matching products found in the database."
        elif hasattr(results, 'empty') and results.empty:
            return "No matching products found in the database."
        elif not len(results):
             return "No matching products found in the database."
        # Limit the results to avoid huge tokens
        if isinstance(results, pd.DataFrame):
            return results.head(5).to_json(orient="records")
        return json.dumps(results[:5], default=str)
    except Exception as e:
        return f"Error searching Cosmos DB: {str(e)}"

def search_web(query):
    """Searches the internet for product details, prices, distributors, and suppliers."""
    try:
        results = DDGS().text(query, max_results=5)
        if not results:
            return "No web results found."
        return json.dumps(list(results))
    except Exception as e:
        return f"Error searching the web: {str(e)}"

def run_personal_assistant(username, user_prompt, files_text="", chat_history=None, is_admin_user=False):
    if chat_history is None:
        chat_history = []
    
    print(f"[PA] run_personal_assistant called. Username={username}, History length={len(chat_history)}, Files={'yes' if files_text else 'no'}", flush=True)
    if chat_history:
        print(f"[PA] Last history message: role={chat_history[-1].get('role')}, content={str(chat_history[-1].get('content'))[:80]}...", flush=True)
    else:
        print("[PA] No chat history — this is a fresh conversation turn.", flush=True)
        
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    time_context = f"\n\nCURRENT DATE AND TIME: {current_time}. Use this to answer any time-related queries accurately."

    if is_admin_user:
        system_prompt = f"""You are a helpful, very efficient, and fast AI personal assistant for {username} (ADMIN).
You are an ADMIN, which means you have access to data of ALL users. 
You can help the user with queries, fetch ALL tasks of EVERYONE or specific users, analyze files, check prices in the local Cosmos DB, and search online.

IMPORTANT INSTRUCTIONS FOR PRICES:
If the user asks about the price or details of a product, you MUST FIRST use `search_cosmos_db` to check the local database. 
If it is not in the database, or if you need to provide online alternatives/links/competitors, you MUST use `search_web` to find online prices and links.
Do not hallucinate prices or links; clearly cite from web results if used.
Return markdown formatting.
{time_context}
"""
    else:
        system_prompt = f"""You are a helpful, very efficient, and fast AI personal assistant for {username}.
Only data belonging to the current user ({username}) is accessible. 
You can help the user with their queries, assist them in tasks, fetch their tasks, analyze files, check prices in the local Cosmos DB, and search online for distributors, suppliers, or online prices.

IMPORTANT INSTRUCTIONS FOR PRICES:
If the user asks about the price or details of a product, you MUST FIRST use `search_cosmos_db` to check the local database. 
If it is not in the database, or if you need to provide online alternatives/links/competitors, you MUST use `search_web` to find online prices and links.
Do not hallucinate prices or links; clearly cite from web results if used.
Return markdown formatting.
{time_context}
"""

    if files_text:
        print(f"[PA] File context added ({len(files_text)} chars)", flush=True)
        system_prompt += f"\nFile Contents Context (user uploaded these for you to analyze):\n{files_text}\n"

    messages = [{"role": "system", "content": system_prompt}]
    
    # Append chat history — only role and content (no timestamp fields)
    for msg in chat_history:
        cleaned = {"role": msg["role"], "content": msg["content"]}
        messages.append(cleaned)
        
    messages.append({"role": "user", "content": user_prompt})
    print(f"[PA] Total messages sent to OpenAI: {len(messages)} (1 system + {len(chat_history)} history + 1 user)", flush=True)

    tools = [
        {
            "type": "function",
            "function": {
                "name": "get_user_tasks",
                "description": "Get tasks from SharePoint. If admin, can optionally specify a target_username or 'all'. Otherwise gets the current user's tasks.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "username": {
                            "type": "string",
                            "description": "The current username."
                        },
                        "target_username": {
                            "type": "string",
                            "description": "(Optional) If admin, the specific user's tasks to fetch. Leave empty or pass 'all' to get all tasks."
                        }
                    },
                    "required": ["username"]
                }
            }
        },
        {
            "type": "function",
            "function": {
                "name": "search_cosmos_db",
                "description": "Search the local Cosmos database for product prices and details.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "query": {
                            "type": "string",
                            "description": "The product name to search for in our database."
                        }
                    },
                    "required": ["query"]
                }
            }
        },
        {
            "type": "function",
            "function": {
                "name": "search_web",
                "description": "Search the internet for product details, prices, distributors, and suppliers.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "query": {
                            "type": "string",
                            "description": "The query to search the web for."
                        }
                    },
                    "required": ["query"]
                }
            }
        },
        {
            "type": "function",
            "function": {
                "name": "draft_email",
                "description": "Draft an email based on the user's request. Always use this when the user asks to write, draft, or compose an email.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "subject": {
                            "type": "string",
                            "description": "The subject line of the email."
                        },
                        "body": {
                            "type": "string",
                            "description": "The body content of the email. Best formatted in plain text or simple markdown."
                        },
                        "to_recipients": {
                            "type": "string",
                            "description": "The recipient's email address. If unknown, leave empty or use a placeholder."
                        }
                    },
                    "required": ["subject", "body", "to_recipients"]
                }
            }
        }
    ]
    try:
        openai.api_key = os.getenv("OPENAI_API_KEY")
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=messages,
            tools=tools,
            temperature=0.0
        )
        
        response_message = response.choices[0].message

        if response_message.get("tool_calls"):
            # Ensure proper dict format for appending
            if isinstance(response_message, dict):
                 messages.append(response_message)
            else:
                 messages.append(response_message.to_dict())
                 
            for tool_call in response_message.tool_calls:
                function_name = tool_call.function.name
                function_args = json.loads(tool_call.function.arguments)
                
                if function_name == "get_user_tasks":
                    target_username = function_args.get("target_username")
                    function_response = get_user_tasks(username, is_admin_user, target_username)
                elif function_name == "search_cosmos_db":
                    function_response = search_cosmos_db(function_args.get("query"))
                elif function_name == "search_web":
                    function_response = search_web(function_args.get("query"))
                elif function_name == "draft_email":
                    # For draft email, we want to return a specific JSON payload back to the frontend
                    # without calling another OpenAI completion, so we just return the payload directly
                    # prefixed with a special tag so the frontend knows how to parse it.
                    function_response = json.dumps({
                        "type": "email_draft",
                        "subject": function_args.get("subject"),
                        "body": function_args.get("body"),
                        "to": function_args.get("to_recipients")
                    })
                    return f"EMAIL_DRAFT_START\n{function_response}\nEMAIL_DRAFT_END"
                else:
                    function_response = "Unknown function"
                    
                messages.append(
                    {
                        "tool_call_id": tool_call.id,
                        "role": "tool",
                        "name": function_name,
                        "content": str(function_response),
                    }
                )
            
            # Second call
            second_response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=messages,
                temperature=0.0
            )
            return second_response.choices[0].message.content
        else:
            return response_message.content

    except Exception as e:
        import traceback
        return f"Error in assistant processing: {str(e)}\n{traceback.format_exc()}"
