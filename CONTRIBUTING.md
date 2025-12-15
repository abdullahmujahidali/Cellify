# Contributing to Cellify

Thank you for your interest in contributing to Cellify! This document provides guidelines and information for contributors.

## Development Setup

1. Clone the repository:

   ```bash
   git clone https://github.com/abdullahmujahidali/Cellify.git
   cd cellify
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

3. Run tests:

   ```bash
   npm test
   ```

4. Build the project:

   ```bash
   npm run build
   ```

## Project Structure

```
cellify/
├── src/
│   ├── core/           # Core classes (Workbook, Sheet, Cell)
│   ├── types/          # TypeScript type definitions
│   ├── formulas/       # Formula parser and evaluator (planned)
│   └── formats/        # Import/export (Excel, CSV) (planned)
├── tests/              # Test files
├── docs/
│   └── decisions/      # Architecture Decision Records
└── package.json
```

## Code Style

- We use TypeScript with strict mode enabled
- Run `npm run lint` to check for type errors
- Write tests for new functionality
- Keep functions focused and well-documented

## Pull Request Process

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass (`npm test`)
6. Commit your changes with a descriptive message
7. Push to your fork
8. Open a Pull Request

## Architecture Decisions

When making significant architectural decisions, please document them in the `docs/decisions/` directory using the ADR format. See existing ADRs for examples.

## Reporting Issues

When reporting issues, please include:

- A clear description of the problem
- Steps to reproduce
- Expected vs actual behavior
- Your environment (Node version, OS)

## Questions?

Feel free to open an issue for questions or discussions about the project.
