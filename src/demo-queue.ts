import { ArrayQueue } from "./structures/queue.ts";

export function demoQueue(): void {
  console.log("=== Завдання 2: Черга на основі масиву ===\n");
  const queue = new ArrayQueue();
  const values = [10, 20, 30, 40, 50];

  console.log("ENQUEUE 5 елементів:");
  for (const v of values) {
    queue.enqueue(v);
    console.log(`  enqueue(${v})`);
  }
  queue.print();

  console.log("\nDEQUEUE 2 елементи:");
  console.log(`  dequeue() → ${queue.dequeue()}`);
  console.log(`  dequeue() → ${queue.dequeue()}`);

  console.log("\nЗалишок черги:");
  queue.print();
}
