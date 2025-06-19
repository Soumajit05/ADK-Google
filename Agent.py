from google.adk.agents import LlmAgent
from google.adk.models.lite_llm import LiteLlm
from tools import get_wikipedia_content, create_dynamic_presentation



PROVIDER = "google"

if PROVIDER == "litellm":
    model=LiteLlm(model="anthropic/claude-3")

elif PROVIDER == "vllm":
    api_base_url = "http://vllm_url/v1"
    model_name = "provider/model_name"
    model=LiteLlm(
        model=model_name,
        api_base=api_base_url,
        api_key="12345"
    )

elif PROVIDER =="google":
    model="gemini-2.0-flash-exp"


get_wikipedia_agent = LlmAgent(
    name="wikipedia_agent",
    description="Agent to search any Wikipedia page and return the complex content.",
    instruction="You are an agent that can search Wikipedia pages and return the complete content by using the given tool.",
    model=model,
    tools=[get_wikipedia_content],
    output_key="data"
)

presentation_agent = LlmAgent(
    name = ("generate_ppt_agent"),
    description=(
        "Generates a PowerPoint Presentatio"
    ),
    instruction=(),
    model=model,
    tools=[create_dynamic_presentation],
)

root_agent = LlmAgent(
    name="coordinator_agent",
    description=(
        "agent to recieve wikipedia content and generate a PowerPoint based on that content"
    ),
    model=model,
    sub_agents=[get_wikipedia_agent, presentation_agent],
)
