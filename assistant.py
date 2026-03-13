import os
import json
import openai
from duckduckgo_search import DDGS
import pandas as pd
from cosmos import search_quotes_by_item
from sharepoint_items import fetch_sharepoint_list

SITE_DOMAIN = "hamdaz1.sharepoint.com"
test_path = "/sites/Test"
test_proposals_list = "testproposals"

def get_my_tasks(username):
    """Fetches tasks for the given username."""
    try:
        tasks = fetch_sharepoint_list(SITE_DOMAIN, test_path, test_proposals_list)
        # Filter logic similar to app.py
        user_tasks = [t for t in tasks if t.get("AssignedTo", "").replace(" ", "") == username]
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

def run_personal_assistant(username, user_prompt, files_text="", chat_history=None):
    if chat_history is None:
        chat_history = []
        
    system_prompt = f"""You are a helpful, very efficient, and fast AI personal assistant for {username}.
Only data belonging to the current user ({username}) is accessible. 
You can help the user with their queries, assist them in tasks, fetch their tasks, analyze files, check prices in the local Cosmos DB, and search online for distributors, suppliers, or online prices.

IMPORTANT INSTRUCTIONS FOR PRICES:
If the user asks about the price or details of a product, you MUST FIRST use `search_cosmos_db` to check the local database. 
If it is not in the database, or if you need to provide online alternatives/links/competitors, you MUST use `search_web` to find online prices and links.
Do not hallucinate prices or links; clearly cite from web results if used.
Return markdown formatting.
"""

    if files_text:
        system_prompt += f"\nFile Contents Context (user uploaded these for you to analyze):\n{files_text}\n"

    messages = [{"role": "system", "content": system_prompt}]
    
    # append chat history
    for msg in chat_history:
        messages.append(msg)
        
    messages.append({"role": "user", "content": user_prompt})

    tools = [
        {
            "type": "function",
            "function": {
                "name": "get_my_tasks",
                "description": "Get the current user's tasks from SharePoint.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "username": {
                            "type": "string",
                            "description": "The current username to fetch tasks for."
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
                
                if function_name == "get_my_tasks":
                    function_response = get_my_tasks(username)
                elif function_name == "search_cosmos_db":
                    function_response = search_cosmos_db(function_args.get("query"))
                elif function_name == "search_web":
                    function_response = search_web(function_args.get("query"))
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
