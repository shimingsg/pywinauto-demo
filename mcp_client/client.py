import asyncio, os
from mcp import ClientSession
from contextlib import AsyncExitStack
from dotenv import load_dotenv
from openai import OpenAI

class MCPClient:
    def __init__(self):
        # self.session = None
        self.exit_stack = AsyncExitStack()
        load_dotenv(dotenv_path="./.env")
        self.ai_api_key = os.getenv("OPENAI_API_KEY")
        self.base_url = os.getenv("BASE_URL")
        self.model = os.getenv("MODEL")
        if not self.ai_api_key:
            raise ValueError("no valid api key")
        
        # self.model='qwen3:4b'
        # self.base_url ='http://localhost:11434/v1'
        # self.ai_api_key = 'ollama'
        
        print(self.ai_api_key)
        print(self.base_url)
        print(self.model)
        
        self.client = OpenAI(api_key=self.ai_api_key, base_url=self.base_url)
    
    async def process_query(self, query:str)->str:
        # messages = [{"role":"system", "content":"you're ai agent, answer qustions from user."},
        #             {"role":"user", "content":query}]
        
        prompt = [{"role":"user", "content":query}]
        
        try:
            response = await asyncio.get_event_loop().run_in_executor(
                None,
                lambda:self.client.chat.completions.create(
                    model=self.model,
                    messages=prompt
                )
            )
            # print(response)
            return response.choices[0].message.content
        except Exception as e:
            return f"Error occurred when calling API: {str(e)}"
    
    async def connect_to_mock_server(self):
        print('MCP client initialized without server')
        
    async def chat_loop(self):
        print('MCP clinet start! Please input "q" to exit')
        while(True):
            try:
                query = input("\nQuery: ").strip()
                if query.lower() == 'q':
                    break
                response = await self.process_query(query)
                print(f"\n [Mock Response]: {response}")
            except Exception as e:
                print(f"\n error occurred: {str(e)}")

    async def cleanup(self):
        await self.exit_stack.aclose()
        
async def main():
    client = MCPClient()
    try:
        await client.connect_to_mock_server()
        await client.chat_loop()
    except Exception as e:
        pass
    finally:
        await client.cleanup()
        
if __name__ == "__main__":
    asyncio.run(main())