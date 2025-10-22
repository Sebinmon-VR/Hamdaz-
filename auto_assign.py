
import openai
from dotenv import load_dotenv
import os
import json
load_dotenv(override=True)

openai.api_key = os.getenv("OPENAI_API_KEY")

def generate_priority_scores_for_all_users(all_users_data):
    """Generates priority for all users in a single call."""
    try:
        users_text = ""
        for u in all_users_data:
            users_text += (
                f"Username: {u['User']}\n"
                f"Total tasks: {u['TotalTasks']}\n"
                f"Completed tasks: {u['CompletedTasksCount']}\n"
                f"Pending tasks: {u['OngoingTasksCount']}\n"
                f"Missed tasks: {u['MissedTasksCount']}\n"
                f"Orders received: {u['OrdersReceived']}\n"
                f"Last assigned date: {u['LastAssignedDate']}\n\n"
            )

        prompt = f"""
You are a task management AI. For the users below, assign a PRIORITY SCORE (1–10) where 1 = most available, 10 = very busy. Consider all users together. 
Respond ONLY with valid JSON: keys = usernames, values = integers (1–10). No text, no commentary.

Users:
{users_text}
        """

        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an AI assistant for task management."},
                {"role": "user", "content": prompt}
            ],
            temperature=0,
            max_tokens=500
        )

        text = response.choices[0].message.content.strip()

        # Parse JSON safely
        try:
            return json.loads(text)
        except Exception as e:
            print("[ERROR] OpenAI output is not valid JSON:", text)
            return {u["User"]: 5 for u in all_users_data}  # fallback

    except Exception as e:
        print("[ERROR] OpenAI call failed:", e)
        return {u["User"]: 5 for u in all_users_data}
