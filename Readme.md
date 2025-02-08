## What is it?

A [Model Context Protocol](https://modelcontextprotocol.io/) Server running over SSE

## What it offers?

Tools for LLM to create PowerPoint presentations as pptx files. The tools can create title layout slides, section layout slides and text content layout slides.

## What do I need?

MCP Client, such is Claude Desktop or [LibreChat](https://github.com/danny-avila/LibreChat)

## How to run this?

Using Docker with precompiled image as per docker-compose.yml. App is listening on port 8958.

### Upload strategy

You must specify upload strategy in env. variable. It may be either local (in such case, a mount of the /app/output folder to host folder is required) or S3 (in such case, AWS credentials and S3 bucket info is required) - see the template docker-compose.yml file

### Custom templates

You may use custom templates so the tool creates the presentation e.g. on your company style slides. In such case, a directory containing "template_4_3.pptx" and "template_16_9.pptx" must be mounted to "/app/templates/" (see docker-compose.yml). In your template, you must ensure the proper position of the slide layout templates in your master slides. The title slide shall be third, title and content slide shall be fifth and section slide shall be eight.

## How to add to LibreChat

In your librechat.yaml file, add the following section:

```yaml
mcpServers:
  media-creator:
    type: sse # type can optionally be omitted
    url: URL of your docker container # e.g. http://localhost:8957/sse
```

## How to use in LibreChat

After the server is added to LibreChat as per above, restart LibreChat to connect to MCP server and discover tools. Then, create an agent and add the respective tools to agent.

When the agent is created, you may ask the agent to create presentation on any topic.

Template system instructions for agent are in instructions_template.md

## Contributions

Further contributions are welcomed!