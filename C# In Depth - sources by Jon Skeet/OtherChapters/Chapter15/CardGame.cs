using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;

namespace Chapter15
{
    class Card
    {
        public const int FullDeckSize = 10;
        // Insert relevant details here

        public static IEnumerable<Card> CreateFullDeck()
        {
            for (int i = 0; i < FullDeckSize; i++)
            {
                yield return new Card();
            }
        }
    }

    class Player
    {
        private readonly List<Card> cards = new List<Card>();

        public int CardCount { get { return cards.Count; } }

        public void AddCard(Card card)
        {
            cards.Add(card);
        }

        public Card RemoveRandomCard()
        {
            Contract.Requires(CardCount > 0);
            Card ret = cards[0]; // TODO: Randomness!
            cards.RemoveAt(0); 
            return ret;
        }
    }

    sealed class CardGame
    {
        private readonly Stack<Card> deck = new Stack<Card>(Card.CreateFullDeck());
        private readonly Stack<Card> discardPile = new Stack<Card>();
        private readonly List<Player> players = new List<Player>();

        public void DealCard(Player player)
        {
            if (deck.Count == 0)
            {
                RecycleDiscards();
            }
            Card card = deck.Pop();
            player.AddCard(card);
        }

        private void RecycleDiscards()
        {
            List<Card> temp = new List<Card>(discardPile);
            // TODO: Code to shuffle temporary list
            // Oh for a PushAll() or similar...
            foreach (Card card in temp)
            {
                deck.Push(card);
            }
        }

        public void ForceDiscard(Player player)
        {
            Card card = player.RemoveRandomCard();
            discardPile.Push(card);
        }

        [ContractInvariantMethod]
        private void ObjectInvariant()
        {
            Contract.Invariant(deck.Count + discardPile.Count +
                               players.Sum(p => p.CardCount) == Card.FullDeckSize);
        }
    }
}
