/* global Word console */

interface Book {
  name: string;
  isbn: string;
  authors: string[];
  numberOfPages: number;
  publisher: string;
  country: string;
  mediaType: string;
  released: string;
  characters: string[];
}

export async function getBooks(): Promise<Book[]> {
  try {
    const response = await fetch("https://www.anapioficeandfire.com/api/books");

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const books: Book[] = await response.json();
    return books;
  } catch (error) {
    console.error("Error fetching books:", error);
    throw error;
  }
}

/**
 * Inserts book details into the Word document.
 */
export async function insertBooksIntoDocument() {
  try {
    const books = await getBooks();
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph("Books List:", Word.InsertLocation.end);

      books.forEach((book) => {
        body.insertParagraph(
          `📖 ${book.name} - ${book.authors.join(", ")} (${book.publisher})`,
          Word.InsertLocation.end
        );
      });

      await context.sync();
    });
  } catch (error) {
    console.error("Error inserting books:", error);
  }
}

//Get characters

interface Aliases {
  aliases: string[];
}

export async function getAliases(): Promise<Aliases[]> {
  try {
    const response = await fetch("https://www.anapioficeandfire.com/api/characters");

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const characters: Aliases[] = await response.json();
    return characters;
  } catch (error) {
    console.error("Error fetching characters:", error);
    throw error;
  }
}

/**
 * Inserts character details into the Word document.
 */

export async function insertAliasesIntoDocument() {
  try {
    const characters = await getAliases();
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph("Characters List:", Word.InsertLocation.end);

      characters.forEach((character) => {
        body.insertParagraph(`📖 ${character.aliases[0]}`, Word.InsertLocation.end);
      });

      await context.sync();
    });
  } catch (error) {
    console.error("Error inserting characters:", error);
  }
}
