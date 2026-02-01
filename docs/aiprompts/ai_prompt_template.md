# AI Prompt Template

Use this template when creating prompts for AI assistants to implement features in Geoffrey.

## Template Structure

```markdown

Do not commit any files to the GIT

# [Feature Name]

## Status
[ ] Not Started | [ ] In Progress | [x] Implemented

## Summary
One paragraph describing what needs to be done and why.

## Context
- What existing code/patterns should be followed?
- What constraints or requirements apply?
- What is the current behavior vs. desired behavior?

## Requirements
1. Specific requirement one
2. Specific requirement two
3. ...

## Files to Modify
- `path/to/file1.py` - What changes are needed
- `path/to/file2.py` - What changes are needed

## Files to Create
- `path/to/new_file.py` - Purpose and key functions

## Implementation Details
Detailed technical specifications, data structures, function signatures, etc.

## Testing
- What tests need to be written?
- How to verify the implementation works?

## Notes
- Edge cases to consider
- Backwards compatibility concerns
- Future considerations
```

## Best Practices

### Be Specific
- Bad: "Add error handling"
- Good: "Add try/except blocks that capture the exception type, message, and traceback, then return them in a structured dict with keys: exception_type, exception_message, traceback"

### Provide Examples
Show concrete before/after code examples when possible:

```python
# Before
return {"status": "error", "message": str(e)}

# After
return {
    "status": "error",
    "error_details": {
        "exception_type": type(e).__name__,
        "message": str(e),
        "traceback": traceback.format_exc()
    }
}
```

### Specify Data Structures
Define the exact shape of data structures:

```python
{
    "field_name": "type and purpose",
    "nested": {
        "inner_field": "description"
    }
}
```

### List Affected Files
Always list which files need to be created, modified, or deleted.

### Include Test Criteria
Specify how to verify the implementation:
- What tests to run
- What manual verification steps
- What edge cases to check

### Consider Backwards Compatibility
Note when existing behavior must be preserved:
- "The existing return structure must remain unchanged"
- "Scripts not using the new feature should continue to work"

### Keep Scope Focused
- One feature per prompt
- Break large features into smaller prompts
- Each prompt should be completable in a single session

## Anti-Patterns to Avoid

1. **Vague requirements**: "Make it better" vs. "Add field X with format Y"
2. **Missing context**: Not explaining what existing patterns to follow
3. **No success criteria**: Not defining how to verify completion
4. **Scope creep**: Mixing unrelated changes in one prompt
5. **Implementation assumptions**: Assuming the AI knows your codebase patterns

## Example: Good Prompt

```markdown
# Add Retry Logic to API Calls

## Status
[ ] Not Started

## Summary
Add configurable retry logic to the external API calls in `api_client.py` to handle transient failures.

## Context
Currently, API calls fail immediately on any error. We need 3 retries with exponential backoff for 5xx errors only.

## Requirements
1. Retry on 500, 502, 503, 504 status codes only
2. Maximum 3 retry attempts
3. Exponential backoff: 1s, 2s, 4s delays
4. Log each retry attempt at WARNING level
5. Raise the original exception after all retries exhausted

## Files to Modify
- `lib/api_client.py` - Add retry decorator or wrapper

## Implementation Details
Create a `@retry_on_5xx` decorator:
- Parameters: max_retries=3, base_delay=1.0
- Use `time.sleep()` for delays
- Check `response.status_code` for retry decision

## Testing
- Unit test with mocked responses returning 503 then 200
- Unit test verifying no retry on 400 errors
- Unit test verifying exception raised after max retries

## Notes
- Don't retry on 4xx errors (client errors)
- Preserve the original exception type when re-raising
```
