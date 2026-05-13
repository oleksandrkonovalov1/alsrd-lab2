class ListNode {
  value: number;
  next: ListNode | null;
  constructor(value: number) {
    this.value = value;
    this.next = null;
  }
}

export class SinglyLinkedList {
  private head: ListNode | null = null;

  append(value: number): void {
    const node = new ListNode(value);
    if (!this.head) {
      this.head = node;
      return;
    }
    let current = this.head;
    while (current.next) current = current.next;
    current.next = node;
  }

  print(): void {
    const items: number[] = [];
    let current = this.head;
    while (current) {
      items.push(current.value);
      current = current.next;
    }
    console.log(`List: ${items.join(" → ")} → null`);
  }
}
