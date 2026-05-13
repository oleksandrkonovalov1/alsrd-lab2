import { ArrayStack } from "./structures/stack.ts";

export function demoStack(): void {
  console.log("=== Завдання 1: Стек на основі масиву ===\n");
  const stack = new ArrayStack();
  const values = [10, 20, 30, 40, 50];

  console.log("PUSH 5 елементів:");
  for (const v of values) {
    stack.push(v);
    console.log(`  push(${v})`);
  }
  stack.print();

  console.log("\nPOP 2 елементи:");
  console.log(`  pop() → ${stack.pop()}`);
  console.log(`  pop() → ${stack.pop()}`);

  console.log("\nЗалишок стеку:");
  stack.print();
}
