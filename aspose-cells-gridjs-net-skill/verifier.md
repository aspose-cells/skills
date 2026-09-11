## Verifying the docs

The **GridJsSkillVerifier** checks that the generated documentation matches the real compiled assemblies and TypeScript declarations.

```powershell
# 1) Compile every ```csharp``` block with Roslyn
dotnet run --project verifier/GridJsSkillVerifier.csproj -c Release -- --compile

# 2) Reflection scan of server‑side inline API references
dotnet run --project verifier/GridJsSkillVerifier.csproj -c Release

# 3) Client‑side API reference validation against index.d.ts
dotnet run --project verifier/GridJsSkillVerifier.csproj -c Release -- --client
```

- **Step 1** ensures all C# snippets are syntactically correct and compile against the referenced Aspose.Cells.GridJs package.
- **Step 2** uses reflection to enumerate every public type, method, property, field, and enum in the server assembly and compares them to the markdown tables.
- **Step 3** parses the `index.d.ts` file shipped with the npm package and validates that every exported TypeScript interface, class, enum, and literal type appears in the client‑side reference.

If any mismatch is reported, update the corresponding markdown file and re‑run the verifier until it exits with code 0. This guarantees that AI agents receive an accurate, up‑to‑date API surface.
