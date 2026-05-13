import { SinglyLinkedList } from "./structures/linked-list.ts";

export function demoList(): void {
  console.log("=== Завдання 3: Зв'язковий список ===\n");
  const list = new SinglyLinkedList();

  console.log("Створення списку з 5 елементів:");
  for (const v of [10, 20, 30, 40, 50]) {
    list.append(v);
    console.log(`  append(${v})`);
  }

  console.log("\nСписок:");
  list.print();
}
