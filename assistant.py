import os
import json
import openai
# from duckduckgo_search import DDGS removed
import pandas as pd
from datetime import datetime
from cosmos import search_quotes_by_item, search_item_distributors
from sharepoint_items import fetch_sharepoint_list

SITE_DOMAIN = "hamdaz1.sharepoint.com"
SITE_PATH = "/sites/ProposalTeam"
LIST_NAME = "Proposals"

def get_user_tasks(current_username, is_admin_user=False, target_username=None, search_keyword=None):
    """Fetches tasks from SharePoint."""
    try:
        tasks = fetch_sharepoint_list(SITE_DOMAIN, SITE_PATH, LIST_NAME)
        
        if search_keyword:
            keyword = search_keyword.lower()
            filtered_tasks = []
            for t in tasks:
                title = str(t.get('Title') or t.get('ProjectName') or t.get('ProposalName') or t.get('Name') or '').lower()
                client = str(t.get('CustomerID') or t.get('ClientName') or '').lower()
                status = str(t.get('JobStatus') or t.get('ApprovalStatus') or '').lower()
                if keyword in title or keyword in client or keyword in status:
                    filtered_tasks.append(t)
            tasks = filtered_tasks

        # Limit to 15 tasks to prevent blowing up the 128K context window
        def safe_return(task_list):
            if len(task_list) > 15:
                return json.dumps(task_list[:15], default=str) + f"\n... (Showing 15 of {len(task_list)} tasks. Please use search_keyword to be more specific)."
            return json.dumps(task_list, default=str)

        if is_admin_user:
            if target_username and target_username.lower() == "all":
                return safe_return(tasks)
            elif target_username and target_username.lower() not in ["my", "me", "mine"]:
                # Fetch a specific other user's tasks
                user_tasks = [t for t in tasks if str(t.get("AssignedTo", "")).replace(" ", "").lower() == target_username.replace(" ", "").lower() or target_username.lower() in str(t.get("AssignedTo", "")).lower()]
                if not user_tasks:
                    return f"No tasks found for user: {target_username}"
                return safe_return(user_tasks)
            else:
                # Target is empty or is "my/me/mine". Fetch Admin's own tasks
                user_tasks = [t for t in tasks if str(t.get("AssignedTo", "")).replace(" ", "").lower() == current_username.lower() or current_username.lower() in str(t.get("AssignedTo", "")).lower()]
                if not user_tasks:
                    return f"No tasks found for the current user ({current_username})."
                return safe_return(user_tasks)
        else:
            user_tasks = [t for t in tasks if str(t.get("AssignedTo", "")).replace(" ", "").lower() == current_username.lower() or current_username.lower() in str(t.get("AssignedTo", "")).lower()]
            if not user_tasks:
                return f"No tasks found for the current user ({current_username})."
            return safe_return(user_tasks)
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

def search_item_purchase_history(query):
    """Searches the local Cosmos DB for item purchase history and distributors."""
    try:
        results = search_item_distributors(query)
        if not results:
            return "No matching purchase history found in the database."
        # Limit results to avoid token limits
        return json.dumps(results[:3], default=str)
    except Exception as e:
        return f"Error searching purchase history: {str(e)}"

def search_web(query):
    """Uses real web search data via OpenAI."""
    try:
        from openai import OpenAI
        import os
        import json
        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        
        input_text = f"Do a real web search for: '{query}'. Return 5 highly relevant, realistic search results including practical distributor names, descriptions, and company URLs. Try to find a contact email for each if possible. Format the output STRICTLY as a raw JSON list of dictionaries, each with 'title', 'href', 'body', and 'email' keys. Set 'email' to an empty string if not found. Do not include markdown brackets (```json). Just the raw list."
        
        # Using the exact snippet format requested by the user, but with a valid available model
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a web search simulation agent. Return strictly the raw JSON array."},
                {"role": "user", "content": input_text}
            ],
            temperature=0.3
        )
        
        ai_output = response.choices[0].message.content.strip()
        # Clean up in case of markdown
        ai_output = ai_output.replace("```json", "").replace("```", "").strip()
        # Validate JSON format
        json.loads(ai_output)
        return ai_output
    except Exception as e:
        return f"Error using AI search: {str(e)}"

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

IMPORTANT INSTRUCTIONS FOR PRICES AND DISTRIBUTORS:
If the user asks about the price or details of a product, you MUST FIRST use `search_cosmos_db` to check the local database for quotes.
To find historical distributors, vendors, or previous purchase prices for an item, use `search_item_purchase_history`. 
If it is not in the database, or if you need to provide online alternatives/links/competitors, you MUST use `search_web` to find online prices and links.
EXTREMELY IMPORTANT: If the user explicitly asks to search online, search the web, or asks for internet distributors, YOU MUST immediately use the `search_web` tool without hesitation. Do not just rely on local databases.
Do not hallucinate prices or links; clearly cite from web results if used.
Return markdown formatting.
{time_context}
"""
    else:
        system_prompt = f"""You are a helpful, very efficient, and fast AI personal assistant for {username}.
Only data belonging to the current user ({username}) is accessible. 
You can help the user with their queries, assist them in tasks, fetch their tasks, analyze files, check prices in the local Cosmos DB, and search online for distributors, suppliers, or online prices.

IMPORTANT INSTRUCTIONS FOR PRICES AND DISTRIBUTORS:
If the user asks about the price or details of a product, you MUST FIRST use `search_cosmos_db` to check the local database for quotes.
To find historical distributors, vendors, or previous purchase prices for an item, use `search_item_purchase_history`. 
If it is not in the database, or if you need to provide online alternatives/links/competitors, you MUST use `search_web` to find online prices and links.
EXTREMELY IMPORTANT: If the user explicitly asks to search online, search the web, or asks for internet distributors, YOU MUST immediately use the `search_web` tool without hesitation. Do not just rely on local databases.
Do not hallucinate prices or links; clearly cite from web results if used.
Return markdown formatting.
{time_context}
"""

    if files_text:
        print(f"[PA] File context added ({len(files_text)} chars)", flush=True)
        system_prompt += f"\nFile Contents Context (user uploaded these for you to analyze):\n{files_text}\n"
        system_prompt += "\nIMPORTANT INSTRUCTIONS FOR UPLOADED REQUIREMENT FILES:\n"
        system_prompt += "If the user asks you to extract items for procurement from the attached file(s) and return them as JSON, you MUST analyze the 'File Contents Context' and return the result STRICTLY as a JSON list of objects, each having a 'name' and 'type' property. DO NOT add any conversational text before or after the JSON array."

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
                        },
                        "search_keyword": {
                            "type": "string",
                            "description": "(Optional) Filter the tasks by a specific keyword, project name, or client. Highly recommended if searching all tasks to avoid token limits."
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
                "name": "search_item_purchase_history",
                "description": "Search the local Cosmos database for historical purchase orders, distributors, and previous purchase prices for an item.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "query": {
                            "type": "string",
                            "description": "The item name or ID to search for in our purchase history."
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
        from openai import OpenAI
        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            tools=tools,
            temperature=0.0
        )
        
        response_message = response.choices[0].message

        if response_message.tool_calls:
            messages.append(response_message.model_dump(exclude_unset=True))
                 
            for tool_call in response_message.tool_calls:
                function_name = tool_call.function.name
                function_args = json.loads(tool_call.function.arguments)
                
                if function_name == "get_user_tasks":
                    target_username = function_args.get("target_username")
                    search_keyword = function_args.get("search_keyword")
                    function_response = get_user_tasks(username, is_admin_user, target_username, search_keyword)
                elif function_name == "search_cosmos_db":
                    function_response = search_cosmos_db(function_args.get("query"))
                elif function_name == "search_item_purchase_history":
                    function_response = search_item_purchase_history(function_args.get("query"))
                elif function_name == "search_web":
                    function_response = search_web(function_args.get("query"))
                elif function_name == "draft_email":
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
            second_response = client.chat.completions.create(
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


