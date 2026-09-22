--[[
figures.lua - tidy the images produced by Convert-ShapesToPictures.ps1 (pandoc >= 3.0)

  1. Title "shape2png:S001" (link to the convert manifest) -> attribute data-shape="S001".
     A title is rendered as a tooltip and is noise for readers and LLMs.
  2. Alt text "Drawing converted to image. Text: A | B" -> "Text in figure: A; B".
     "|" collides with table syntax and gets escaped as "\|".
  3. Figure whose caption starts with an empty bookmark span ([]{#_Ref123 .anchor}):
     the id moves to the figure, so cross references (#_Ref123) keep working and the
     caption no longer contains nested brackets, which many markdown viewers mis-render.

Usage: pandoc -f docx -t markdown --lua-filter=figures.lua input.docx -o output.md
]]

local ALT_PREFIX = "Drawing converted to image. Text: "
local ALT_EMPTY = "Drawing converted to image"

local function normalize_alt(s)
  if s:sub(1, #ALT_PREFIX) == ALT_PREFIX then
    local body = s:sub(#ALT_PREFIX + 1):gsub(" | ", "; ")
    return "Text in figure: " .. body
  elseif s == ALT_EMPTY then
    return "Figure without text"
  end
  return nil
end

function Image(img)
  local id = img.title:match("^shape2png:(.+)$")
  if not id then
    return nil
  end
  img.title = ""
  img.attributes["data-shape"] = id

  -- Inside a Figure, pandoc moves the docx description to the "alt" attribute and uses
  -- the caption as image content; outside a Figure the description is the content.
  local alt = img.attributes["alt"]
  if alt then
    local new = normalize_alt(alt)
    if new then img.attributes["alt"] = new end
  else
    local new = normalize_alt(pandoc.utils.stringify(img.caption))
    if new then img.caption = { pandoc.Str(new) } end
  end
  return img
end

function Figure(fig)
  if fig.identifier ~= "" then
    return nil
  end
  local anchor
  fig.caption.long = fig.caption.long:walk({
    Span = function(span)
      if not anchor and span.classes:includes("anchor") and #span.content == 0 then
        anchor = span.identifier
        return {}
      end
    end,
  })
  if not anchor then
    return nil
  end
  fig.identifier = anchor
  return fig
end
