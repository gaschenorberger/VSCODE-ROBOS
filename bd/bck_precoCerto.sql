PGDMP      1                }            bd_preco_certo    17.2    17.2     �           0    0    ENCODING    ENCODING        SET client_encoding = 'UTF8';
                           false            �           0    0 
   STDSTRINGS 
   STDSTRINGS     (   SET standard_conforming_strings = 'on';
                           false            �           0    0 
   SEARCHPATH 
   SEARCHPATH     8   SELECT pg_catalog.set_config('search_path', '', false);
                           false            �           1262    195423    bd_preco_certo    DATABASE     �   CREATE DATABASE bd_preco_certo WITH TEMPLATE = template0 ENCODING = 'UTF8' LOCALE_PROVIDER = libc LOCALE = 'Portuguese_Brazil.1252';
    DROP DATABASE bd_preco_certo;
                     postgres    false            �            1259    195435    produtos    TABLE     �   CREATE TABLE public.produtos (
    id integer NOT NULL,
    nome_pdt character varying(100) NOT NULL,
    preco_pdt money NOT NULL,
    link_pdt text NOT NULL
);
    DROP TABLE public.produtos;
       public         heap r       postgres    false            �            1259    195434    produtos_id_seq    SEQUENCE     �   CREATE SEQUENCE public.produtos_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;
 &   DROP SEQUENCE public.produtos_id_seq;
       public               postgres    false    218            �           0    0    produtos_id_seq    SEQUENCE OWNED BY     C   ALTER SEQUENCE public.produtos_id_seq OWNED BY public.produtos.id;
          public               postgres    false    217            W           2604    195438    produtos id    DEFAULT     j   ALTER TABLE ONLY public.produtos ALTER COLUMN id SET DEFAULT nextval('public.produtos_id_seq'::regclass);
 :   ALTER TABLE public.produtos ALTER COLUMN id DROP DEFAULT;
       public               postgres    false    218    217    218            �          0    195435    produtos 
   TABLE DATA           E   COPY public.produtos (id, nome_pdt, preco_pdt, link_pdt) FROM stdin;
    public               postgres    false    218          �           0    0    produtos_id_seq    SEQUENCE SET     =   SELECT pg_catalog.setval('public.produtos_id_seq', 5, true);
          public               postgres    false    217            Y           2606    195442    produtos produtos_pkey 
   CONSTRAINT     T   ALTER TABLE ONLY public.produtos
    ADD CONSTRAINT produtos_pkey PRIMARY KEY (id);
 @   ALTER TABLE ONLY public.produtos DROP CONSTRAINT produtos_pkey;
       public                 postgres    false    218            �   �  x���َ�F�����R2��w�T�DjE��6�666E�]f)e�y�<E^,���2e2-�Pu��u>N�ſ�Eu��@�Ë^P�~g�i�󨛓�tqTRԙdET��YG����S�_9�ՁҢ���1Hq��YT($)���! ك_>f���9�z�.4*s�>���c@�9�=)�&w-~��ǋ��}-hE�Z�r���C;65O�C�ѓfG�	X�S�R��6��,�y�n,A֛����dX�<�)�T�=$+��s���A��vU�憒�)s���j����B�Pl9�Jyt��ir,��R�^�9��K��Az�Z�;����N�:��^Maki=�"�3��-���|����9���Qq9vZΦ���n�Ѫ�U�a-�8B�=i��nO��i��ZsJ�)�e�������Gă�W�JT5��VW��e����.��9GQ/A�}}J�@��y�'��Ԭ���^���g�.E����0��`z���/ǹ����~U/=��ع��@�x��mO=�U��#cBr�jGC�n�ώ�y{���yխ����,�����.��g��:.\\�Tj�8�3��i�<���z�N^``�hx���S��3cvY����3 ����cV.����gp��f}�%��x��)**a�O���_��cׇ���K{5q�3ub�z�Vk�����p��ㆍAu��X�n,7syVV��pz4p֞x���	�viO1��~���I�
S'LXEX�G��n��b���� g9�7�.����W�scW娵�C4�rl�!YUl8R�0���{�F�5$I՘�����du��n���a�d7�xu1�m����z�XJ���ZX�#�j���ٹ�ؿ����t�� �5@e���(�OU�����qpOK$8߿��S,�@� �*0����\�W��G!'
0�}	�	_���3��K��U4JQ�ކ�b��qH�N�LI�[f%�n$��4M�V�nB]\܄���e:ݽߍ�E��g]�Qⱸ���
�8�o x�ڢ#�I�����aa+"d>8Q����@�}��]ə����=��SQ����{_���w���8�������rV%���"��s?rV xE�O������;o(���k��C�_^��2�2�d��EX�yY�d�%�?�~ ��@�W�O��%���w��a3¨3'�F?�3��9c:���������,	�,��?і~,�?����I(W�     